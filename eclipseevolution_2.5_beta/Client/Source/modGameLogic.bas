Attribute VB_Name = "modGameLogic"
'*********************************
' _____     _ _
'|  ___|   | (_)
'| |__  ___| |_ _ __  ___  ___
'|  __|/ __| | | '_ \/ __|/ _ \
'| |__| (__| | | |_) \__ \  __/
'\____/\___|_|_| .__/|___/\___|
'              | |
'              |_|
'*********************************
' ECLIPSE EVO CLIENT
'   COMPILING OPTIMIZATIONS:
'    Compiled for small size, due to large-ness of client taking up memory
'    Removed safe pentium checks, 'cause that issue was resolved back in like friggin 1996


Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public snumber As Integer
Public RWINDEX As Long

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
Public SS_WALK_SPEED
Public SS_RUN_SPEED
'Set the variable to your desire,
'32 is a safe and recommended setting

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Used for AlwaysOnTop
Const FLAGS As Long = 3
Const HWND_TOPMOST As Long = -1
Const HWND_NOTOPMOST As Long = -2
Public SetTop As Boolean
Private Declare Function SetWindowPos Lib "user32" (ByVal H As Long, ByVal hb As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal f As Long) As Long

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

' Used for map key open editor
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
Public InSkillEditor As Boolean
Public InQuestEditor As Boolean
Public EditorIndex As Long

' Game fps
Public GameFPS As Long
Public BFPS As Boolean

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

' Used for NPC spawn
Public NPCSpawnNum As Long

' Used for roof tile
Public RoofId As String

Public Wierd As Long
Public AutoLogin As Long

'Used to make sure we have all the data before logging in
Public AllDataReceived As Boolean

' Used for classes
Public ClassesOn As Byte

'Keep track of time
Public Hours As Integer
Public Minutes As Integer
Public Seconds As Integer
Public Gamespeed As Integer

' Font data
Public Font As String
Public fontsize As Byte

'Main sub. Client starts here
Sub Main()
    
    On Error GoTo LoadErr
    ScreenMode = 0
    
    MOUSE_HEIGHT = Trim$(ReadINI("CONFIG", "Mouse_Height", App.Path & "\config.ini"))
MOUSE_WIDTH = Trim$(ReadINI("CONFIG", "Mouse_Width", App.Path & "\config.ini"))


    frmSendGetData.Visible = True
    Call SetStatus("Checking folders...")
    DoEvents
    
    'Check all the DLLs and data files
    Call SystemFileChecker
    
    ' Check if the maps directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\Maps", vbDirectory)) <> "maps" Then
        Call MkDir$(App.Path & "\Maps")
    End If
    If UCase$(Dir$(App.Path & "\GFX", vbDirectory)) <> "GFX" Then
        Call MkDir$(App.Path & "\GFX")
    End If
    If UCase$(Dir$(App.Path & "\GUI", vbDirectory)) <> "GUI" Then
        Call MkDir$(App.Path & "\GUI")
    End If
    If UCase$(Dir$(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Call MkDir$(App.Path & "\Music")
    End If
    If UCase$(Dir$(App.Path & "\SFX", vbDirectory)) <> "SFX" Then
        Call MkDir$(App.Path & "\SFX")
    End If
    If UCase$(Dir$(App.Path & "\Flashs", vbDirectory)) <> "FLASHS" Then
        Call MkDir$(App.Path & "\Flashs")
    End If
    If UCase$(Dir$(App.Path & "\BGS", vbDirectory)) <> "BGS" Then
        Call MkDir$(App.Path & "\BGS")
    End If
    If UCase$(Dir$(App.Path & "\DATA", vbDirectory)) <> "DATA" Then
        Call MkDir$(App.Path & "\Data")
    End If
    
    Call SetStatus("Loading data...")
    
    On Error GoTo configErr
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
        WriteINI "UPDATER", "WebSite", vbNullString, App.Path & "\config.ini"
        WriteINI "IPCONFIG", "IP", "127.0.0.1", App.Path & "\config.ini"
        WriteINI "IPCONFIG", "PORT", 4001, App.Path & "\config.ini"
        WriteINI "CONFIG", "Account", vbNullString, App.Path & "\config.ini"
        WriteINI "CONFIG", "Password", vbNullString, App.Path & "\config.ini"
        WriteINI "CONFIG", "WebSite", vbNullString, App.Path & "\config.ini"
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
    
    GoTo configSuccess
    
configErr:
    Call MsgBox("Error reading from config.ini. Make sure all varialbes are valid! Err: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    End
    
configSuccess:
    'Return error handling to the one specified above
    On Error GoTo 0
    
    If FileExist("News.ini") = False Then
        WriteINI "DATA", "News", "News:*Eclipse has been released", App.Path & "\News.ini"
        WriteINI "DATA", "Desc", "Description:Enter Description here", App.Path & "\News.ini"
        WriteINI "COLOR", "Red", 255, App.Path & "\News.ini"
        WriteINI "COLOR", "Green", 255, App.Path & "\News.ini"
        WriteINI "COLOR", "Blue", 255, App.Path & "\News.ini"
        WriteINI "FONT", "Font", "Arial", App.Path & "\News.ini"
        WriteINI "FONT", "Size", "14", App.Path & "\News.ini"
    End If
    
    Call SetStatus("Loading Colors and font...")
    
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
    
    If Not FileExist("Font.ini") Then
        WriteINI "FONT", "Font", "Lucida Sans", App.Path & "\Font.ini"
        WriteINI "FONT", "Size", 16, App.Path & "\Font.ini"
    End If
    
    GoTo LoadSuccess
    
LoadErr:
    Call MsgBox("Error loading settings in Main(). Please check all your files. ERR: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    End
    
LoadSuccess:
    On Error GoTo Load2Err
    Dim R1 As Long, G1 As Long, B1 As Long
    R1 = val#(ReadINI("CHATBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = val#(ReadINI("CHATBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = val#(ReadINI("CHATBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtChat.BackColor = RGB(R1, G1, B1)
    
    R1 = val#(ReadINI("CHATTEXTBOX", "R", App.Path & "\GUI\Colors.txt"))
    G1 = val#(ReadINI("CHATTEXTBOX", "G", App.Path & "\GUI\Colors.txt"))
    B1 = val#(ReadINI("CHATTEXTBOX", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.txtMyTextBox.BackColor = RGB(R1, G1, B1)
    
    R1 = val#(ReadINI("BACKGROUND", "R", App.Path & "\GUI\Colors.txt"))
    G1 = val#(ReadINI("BACKGROUND", "G", App.Path & "\GUI\Colors.txt"))
    B1 = val#(ReadINI("BACKGROUND", "B", App.Path & "\GUI\Colors.txt"))
    
   ' frmMirage.Picture9.BackColor = RGB(R1, G1, B1)
    frmMirage.picInv3.BackColor = RGB(R1, G1, B1)
    frmMirage.itmDesc.BackColor = RGB(R1, G1, B1)
    frmMirage.picWhosOnline.BackColor = RGB(R1, G1, B1)
    frmMirage.picGuildAdmin.BackColor = RGB(R1, G1, B1)
    frmMirage.Picture1(0).BackColor = RGB(R1, G1, B1)
  '  frmMirage.picEquip.BackColor = RGB(R1, G1, B1)
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
    
    R1 = val#(ReadINI("SPELLLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = val#(ReadINI("SPELLLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = val#(ReadINI("SPELLLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstSpells.BackColor = RGB(R1, G1, B1)
    
    R1 = val#(ReadINI("WHOLIST", "R", App.Path & "\GUI\Colors.txt"))
    G1 = val#(ReadINI("WHOLIST", "G", App.Path & "\GUI\Colors.txt"))
    B1 = val#(ReadINI("WHOLIST", "B", App.Path & "\GUI\Colors.txt"))
    frmMirage.lstOnline.BackColor = RGB(R1, G1, B1)
    
    R1 = val#(ReadINI("NEWCHAR", "R", App.Path & "\GUI\Colors.txt"))
    G1 = val#(ReadINI("NEWCHAR", "G", App.Path & "\GUI\Colors.txt"))
    B1 = val#(ReadINI("NEWCHAR", "B", App.Path & "\GUI\Colors.txt"))
    frmNewChar.optMale.BackColor = RGB(R1, G1, B1)
    frmNewChar.optFemale.BackColor = RGB(R1, G1, B1)
    
    Font = ReadINI("FONT", "Font", App.Path & "\Font.ini")
    fontsize = val(ReadINI("FONT", "Size", App.Path & "\Font.ini"))
    
    If Font + vbNullString = vbNullString Then
        Font = "Lucida Console"
    End If
    If fontsize + 0 = 0 Then
        fontsize = 16
    End If
    GoTo Load2Success
    
Load2Err:
    Call MsgBox("Error loading colors in Main(). Please check your colors.txt and font.ini files. ERR:" & Err.Number & ", DESC: " & Err.Description, vbCritical)
    End
    
Load2Success:
    Call SetStatus("Checking status...")
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
    InQuestEditor = False
    
    Call SetStatus("Initializing TCP Settings...")
    DoEvents
    
    Call TcpInit
    
    frmCredits.Label15.Caption = "  This game was built using Eclipse Evolution  www.touchofdeathproductions.com"
    frmCredits.Label15.Height = 33
    frmCredits.Label15.Width = 371
    frmCredits.Label15.Left = 16
    frmCredits.Label15.Top = 96
    
    
    Screen_RESIZED = 0
    
    If ReadINI("CONFIG", "Res", App.Path & "\config.ini") & vbNullString = vbNullString Then
        frmMirage.ScrlResolution.Value = 1
    Else
        frmMirage.ScrlResolution.Value = Int(ReadINI("CONFIG", "Res", App.Path & "\config.ini"))
    End If
    
    'We have not yet recieved any data, so:
    AllDataReceived = False
    frmMainMenu.Visible = True
    Unload frmSendGetData
    
    
End Sub

Function ExactFileExist(ByVal FileName As String) As Boolean
    If Dir$(FileName) = vbNullString Then
        ExactFileExist = False
    Else
        ExactFileExist = True
    End If
End Function

Function ExactCopyFile(ByVal Source As String, ByVal Destination As String)
    On Error GoTo CopyError
    
    FileCopy Source, Destination
    ExactCopyFile = True
    
CopyError:     MsgBox Source & " is missing!"
    
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


Sub SystemFileChecker()
    On Error GoTo RegErr
'    FILE LIST TO BE COPIED AND REGISTERED WITH CODE
    
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
    Call registerfilechecker("scrrun.dll")
    Call registerfilechecker("SSubTmr6.dll")
    Call registerfilechecker("SSubTmr.dlm")
    Exit Sub
    
RegErr:
    'Error handler
    Call MsgBox("Error checking and/or registering a file in data. Make sure all the files are there, or try registering them manually with regsvr32. Otherwise, the client may not work!", vbCritical)
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
    On Error Resume Next
    frmSendGetData.lblStatus.Caption = Caption
End Sub

Sub MenuState(ByVal State As Long)
    Connucted = True
    frmSendGetData.Visible = True
    Call SetStatus("Connecting to server...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            Unload frmNewAccount
            If ConnectToServer = True Then
                Call StopBGM
                Call SetStatus("Sending new account information...")
                If frmNewAccount.txtEmail.Text = vbNullString And frmNewAccount.txtEmail.Visible = False Then
                    Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text, "NOMAIL")
                Else
                    Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text, frmNewAccount.txtEmail.Text)
                End If
                '...why?
                'frmMirage.Socket.Close
                frmSendGetData.Visible = True
                Exit Sub
            End If
            
        Case MENU_STATE_DELACCOUNT
            Unload frmDeleteAccount
            If ConnectToServer = True Then
                Call SetStatus("Sending account deletion request ...")
                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
            End If
        
        Case MENU_STATE_LOGIN
            Unload frmLogin
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If
            
        Case MENU_STATE_AUTO_LOGIN
            Unload frmMainMenu
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If
        
        Case MENU_STATE_NEWCHAR
            
            If ConnectToServer = True Then
                    If Spritesize = 1 Then
                    frmNewChar.Picture4.Top = frmNewChar.Picture4.Top - 32
                    frmNewChar.Picture4.Height = 69
                    frmNewChar.picPic.Height = 65
                    End If
                If 0 + customplayers <> 0 Then
                    frmNewChar.HScroll1.Visible = True
                    frmNewChar.HScroll2.Visible = True
                    frmNewChar.HScroll3.Visible = True
                End If
                Call SetStatus("Connected, getting available classes...")
                Call SendGetClasses
            End If
            Unload frmChars
        Case MENU_STATE_ADDCHAR
            
            If ConnectToServer = True Then
                Call SetStatus("Sending character addition data...")
                If frmNewChar.optMale.Value = True Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1, frmNewChar.HScroll1.Value, frmNewChar.HScroll2.Value, frmNewChar.HScroll3.Value)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1, frmNewChar.HScroll1.Value, frmNewChar.HScroll2.Value, frmNewChar.HScroll3.Value)
                End If

            End If
            Unload frmNewChar
        Case MENU_STATE_DELCHAR
            
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character deletion request...")
                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
            End If
            Unload frmChars
        Case MENU_STATE_USECHAR
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending char data...")
                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
            Unload frmChars
    End Select
    
    If Not IsConnected And Connucted = True Then
        frmMainMenu.Visible = True
        Unload frmSendGetData
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub
Sub GameInit()
    Call StopBGM
    Call InitDirectX
    
'    Alpha blending - unused
'    BF.BlendOp = AC_SRC_OVER
'    BF.BlendFlags = 0
'    BF.AlphaFormat = 0
    
    frmMirage.Visible = True
    Unload frmSendGetData
    Call frmMirage.SetFocus
    
End Sub

Sub GiveItems()
Dim i As Long
If Player(MyIndex).Inv(RCIINDEX + 1).num <> 0 Then
    If item(Player(MyIndex).Inv(RCIINDEX + 1).num).Stackable = 1 Or item(Player(MyIndex).Inv(RCIINDEX + 1).num).Type = ITEM_TYPE_CURRENCY Then
        i = val(InputBox("Give how many?", "Gift", 0))
        If IsNumeric(i) Then
            If i <= Player(MyIndex).Inv(RCIINDEX + 1).Value And i <> 0 Then
                GIFTAMT = i
            ElseIf i = 0 Then
                Call MsgBox("Invalid Amount")
                Exit Sub
            Else
                Call MsgBox("You don't have that many")
                Exit Sub
            End If
        Else
            Call MsgBox("Invalid Amount")
            Exit Sub
        End If
    Else
        GIFTAMT = 1
    End If
    GIFTTO = InputBox("Give to whom?", "Gift")
    Call SendData("sendinggiftto" & SEP_CHAR & GIFTTO & SEP_CHAR & GIFTAMT & SEP_CHAR & Player(MyIndex).Inv(RCIINDEX + 1).num)
    DoEvents
End If
End Sub
Sub GameLoop()
Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim z As Long
Dim connectionLost As Boolean

' Game Loop error handler
 ' On Error GoTo GameErr

'Update the stats bars
'  Check for divide by 0 error
If GetPlayerMaxHP(MyIndex) > 0 Then
    frmMirage.shpHP.Width = ((GetPlayerHP(MyIndex)) / (GetPlayerMaxHP(MyIndex))) * 150
    frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
End If
'   Check for divide by 0 error
If GetPlayerMaxMP(MyIndex) > 0 Then
    frmMirage.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
    frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
End If
            
' Set the focus To the main form since only focussed objects may Set the focus
frmMirage.SetFocus

' Set the focus
frmMirage.picUber.SetFocus
'frmMirage.picScreen.SetFocus

' Set font
Call SetFont(Font, fontsize)

' Used for calculating fps
TickFPS = GetTickCount
FPS = 0

connectionLost = False

'           ********************
'*******************************************
'* ECLIPSE EVOLUTION MAIN GAME LOOP BEGIN  *
'*******************************************
'           ********************
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
If Not IsConnected Then
    InGame = False
    connectionLost = True
End If

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

NewX = 10
NewY = 7

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
                Call BitBlt(frmMirage.picInv(Q).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(Qq).Pic - Int(item(Qq).Pic / 6) * 6) * PIC_X, Int(item(Qq).Pic / 6) * PIC_Y, SRCCOPY)
            End If
        End If
    Next Q
End If

NewPlayerY = Player(MyIndex).Y - NewY
NewPlayerX = Player(MyIndex).X - NewX

NewX = NewX * PIC_X
NewY = NewY * PIC_Y

NewXOffset = Player(MyIndex).xOffset
NewYOffset = Player(MyIndex).yOffset

If Player(MyIndex).Y - 7 < 1 Then
    NewY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).yOffset
    NewYOffset = 0
    NewPlayerY = 0
    If Player(MyIndex).Y = 7 And Player(MyIndex).Dir = DIR_UP Then
        NewPlayerY = Player(MyIndex).Y - 7
        NewY = 7 * PIC_Y
        NewYOffset = Player(MyIndex).yOffset
    End If
    ElseIf Player(MyIndex).Y + 9 > MAX_MAPY + 1 Then
        NewY = (Player(MyIndex).Y - 16) * PIC_Y + Player(MyIndex).yOffset
        NewYOffset = 0
        NewPlayerY = MAX_MAPY - 14
    If Player(MyIndex).Y = 23 And Player(MyIndex).Dir = DIR_DOWN Then
        NewPlayerY = Player(MyIndex).Y - 7
        NewY = 7 * PIC_Y
        NewYOffset = Player(MyIndex).yOffset
    End If
End If

If Player(MyIndex).X - 10 < 1 Then
    NewX = Player(MyIndex).X * PIC_X + Player(MyIndex).xOffset
    NewXOffset = 0
    NewPlayerX = 0
    If Player(MyIndex).X = 10 And Player(MyIndex).Dir = DIR_LEFT Then
        NewPlayerX = Player(MyIndex).X - 10
        NewX = 10 * PIC_X
        NewXOffset = Player(MyIndex).xOffset
    End If
    ElseIf Player(MyIndex).X + 11 > MAX_MAPX + 1 Then
        NewX = (Player(MyIndex).X - 11) * PIC_X + Player(MyIndex).xOffset
        NewXOffset = 0
        NewPlayerX = MAX_MAPX - 19
        If Player(MyIndex).X = 21 And Player(MyIndex).Dir = DIR_RIGHT Then
        NewPlayerX = Player(MyIndex).X - 10
        NewX = 10 * PIC_X
        NewXOffset = Player(MyIndex).xOffset
    End If
End If

sx = 32
If MAX_MAPX = 19 Then
    NewX = Player(MyIndex).X * PIC_X + Player(MyIndex).xOffset
    NewXOffset = 0
    NewPlayerX = 0
    NewY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).yOffset
    NewYOffset = 0
    NewPlayerY = 0
    sx = 0
End If

' Clear the back buffer
ClearBackBuffer

' Blit out tiles layers ground/anim1/anim2
For Y = 0 To MAX_MAPY
    For X = 0 To MAX_MAPX
        Call BltTile(X, Y)
    Next X
Next Y

If ScreenMode = 0 Then
' Blit out the items
For i = 1 To MAX_MAP_ITEMS
    If MapItem(i).num > 0 Then
        Call BltItem(i)
    End If
Next i

If frmMirage.chknpcbar.Value = Checked Then
    ' Blit out NPC hp bars
    For i = 1 To MAX_MAP_NPCS
        Call BltNpcBars(i)
    Next i
End If
X = 0
For i = 1 To MAX_PARTY_MEMBERS
If Player(MyIndex).Party.Member(i) <> MyIndex And Player(MyIndex).Party.Member(i) <> 0 Then X = 1
Next i
If X = 0 Then
For i = 1 To MAX_PARTY_MEMBERS
Player(MyIndex).Party.Member(i) = 0
Next i
End If
If frmMirage.chkplayerbar.Value = Checked Then
    ' Blit players bar
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
        If i <> MyIndex Then
            For z = 1 To MAX_PARTY_MEMBERS
                If Player(MyIndex).Party.Member(z) = i Then
          Call SendData("getplayerhp" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
            Call BltPlayerBars(i)
                    
                End If
            Next z
          End If
        End If
    Next i
    Call BltPlayerBars(MyIndex)
End If

' Blit out the sprite change attribute
For Y = 0 To MAX_MAPY
    For X = 0 To MAX_MAPX
        Call BltSpriteChange(X, Y)
    Next X
Next Y

' Blit out arrows
For i = 1 To MAX_PLAYERS
If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
Call BltArrow(i)
End If
Next i

' Blit out grapple
For i = 1 To MAX_PLAYERS
    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
        Call Bltgrapple(i)
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

If Spritesize >= 1 Then
    ' Blit out players top
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

' Check for roof tiles
For Y = 0 To MAX_MAPY
    For X = 0 To MAX_MAPX
        If Not IsTileRoof(X, Y) Then
            Call BltFringe2Tile(X, Y)
        End If
    Next X
Next Y

'Draw 'level up!' text
For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If Player(i).LevelUpT + 3000 > GetTickCount Then
            rec.Top = Int(32 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (32 - Int(32 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + 96
            
            If i = MyIndex Then
                X = NewX + sx
                Y = NewY + sx
                Call DD_BackBuffer.BltFast(X - 32, Y - 10 - Player(i).LevelUp, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                X = GetPlayerX(i) * PIC_X + sx + Player(i).xOffset
                Y = GetPlayerY(i) * PIC_Y + sx + Player(i).yOffset
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - 32 - NewXOffset, Y - (NewPlayerY * PIC_Y) - 10 - Player(i).LevelUp - NewYOffset, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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

'Draw weather and night
If GettingMap = False Then
    If Wierd = 1 Then
        Call WierdNight
    Else
    If GameTime = TIME_NIGHT And Map(GetPlayerMap(MyIndex)).Indoors = 0 And InEditor = False And 0 + Map(GetPlayerMap(MyIndex)).lights = 0 Then
        Call Night
    End If
    If frmMapEditor.mnuDayNight.Checked = True And InEditor = True Then
        Call Night
    End If
    If Map(GetPlayerMap(MyIndex)).Indoors = 0 And Map(GetPlayerMap(MyIndex)).Weather = 0 Then Call BltWeather
    End If
End If

If Map(GetPlayerMap(MyIndex)).Weather <> 0 Then Call BltMapWeather

'Blit cannon (unused)
'If CanonUsed <> 0 Then Call BltCanon

If (InEditor = True Or InHouseEditor = True) And ReadINI("CONFIG", "MapGrid", App.Path & "\config.ini") = 1 Then
For Y = 0 To MAX_MAPY
For X = 0 To MAX_MAPX
Call BltTile2(X * 32, Y * 32, 0)
Next X
Next Y
End If
End If

' Lock the backbuffer so we can draw text and names
TexthDC = DD_BackBuffer.GetDC
If GettingMap = False Then
If ScreenMode = 0 Then
If frmMirage.chknpcdamage.Value = 1 Then
If frmMirage.chkplayername.Value = 0 Then
If GetTickCount < NPCDmgTime + 2000 Then
Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BrightRed))
End If
Else
If GetPlayerGuild(MyIndex) <> vbNullString Then
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

If frmMirage.chkplayerdamage.Value = 1 Then
If NPCWho > 0 Then
If MapNpc(NPCWho).num > 0 Then
If frmMirage.chknpcname.Value = 0 Then
If Npc(MapNpc(NPCWho).num).Big = 0 Then
If GetTickCount < DmgTime + 2000 Then
Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(White))
End If
Else
If GetTickCount < DmgTime + 2000 Then
Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 47 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(White))
End If
End If
Else
If Npc(MapNpc(NPCWho).num).Big = 0 Then
If GetTickCount < DmgTime + 2000 Then
Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(White))
End If
Else
If GetTickCount < DmgTime + 2000 Then
Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 57 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(White))
End If
End If
End If
iii = iii + 1
End If
End If
End If

If frmMirage.chkplayername.Value = 1 Then
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
If Bubble(i).Text <> vbNullString Then
Call BltPlayerText(i)
End If

If GetTickCount() > Bubble(i).Created + DISPLAY_BUBBLE_TIME Then
Bubble(i).Text = vbNullString
End If
End If
Next i
End If

' scriptbubble stuffs
i = MyIndex
For z = 1 To MAX_BUBBLES
    If IsPlaying(i) And GetPlayerMap(i) = ScriptBubble(z).Map Then
        
            If ScriptBubble(z).Text <> vbNullString Then
            Call Bltscriptbubble(z, ScriptBubble(z).Map, ScriptBubble(z).X, ScriptBubble(z).Y, ScriptBubble(z).Colour)
            End If

            If GetTickCount() > ScriptBubble(z).Created + DISPLAY_BUBBLE_TIME Then
            ScriptBubble(z).Text = vbNullString
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


End If

' Blit out attribs if in editor
If InEditor Or InHouseEditor Then

For Y = 0 To MAX_MAPY
For X = 0 To MAX_MAPX
With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
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
If .Type = TILE_TYPE_NPC_SPAWN Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "NPC", QBColor(BrightGreen))
If .Type = TILE_TYPE_HOUSE Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PH", QBColor(Yellow))
If .light > 0 Then Call DrawText(TexthDC, X * PIC_X + sx + 18 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 14 - (NewPlayerY * PIC_Y) - NewYOffset, "L", QBColor(Yellow))
If .Type = TILE_TYPE_BANK Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "BANK", QBColor(BrightRed))
If .Type = TILE_TYPE_CANON Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CA", QBColor(Yellow))
If .Type = TILE_TYPE_SKILL Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SK", QBColor(Yellow))
If .Type = TILE_TYPE_GUILDBLOCK Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "GB", QBColor(Magenta))
If .Type = TILE_TYPE_HOOKSHOT Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "GS", QBColor(White))
If .Type = TILE_TYPE_WALKTHRU Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "WT", QBColor(Red))
If .Type = TILE_TYPE_ROOF Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "RF", QBColor(Red))
If .Type = TILE_TYPE_ROOFBLOCK Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "RFB", QBColor(BrightRed))
If .Type = TILE_TYPE_ONCLICK Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "OC", QBColor(White))
If .Type = TILE_TYPE_LOWER_STAT Then Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "-S", QBColor(BrightRed))
End With
Next X
Next Y
End If

' Blit the text they are putting in
'MyText = frmMirage.txtMyTextBox.Text
'frmMirage.txtMyTextBox.Text = MyText

'If Len(MyText) > 4 Then
'frmMirage.txtMyTextBox.SelStart = Len(frmMirage.txtMyTextBox.Text) + 1
'End If

' draw FPS
            If BFPS = True Then
                Call DrawText(TexthDC, Int((MAX_MAPX - 1) * PIC_X), 1, Trim$("FPS: " & GameFPS), QBColor(Yellow))
            End If


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
Call DrawText(TexthDC, 1 + sx, BattlePMsg(i).Y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattlePMsg(i).Msg), QBColor(BattlePMsg(i).color))
Else
BattlePMsg(i).Done = 0
End If
End If

If BattleMMsg(i).Index > 0 Then
If BattleMMsg(i).Time + 7000 > GetTickCount Then
Call DrawText(TexthDC, (frmMirage.picScreen.Width - (Len(BattleMMsg(i).Msg) * 8)) + sx, BattleMMsg(i).Y + frmMirage.picScreen.Height - 15 + sx, Trim$(BattleMMsg(i).Msg), QBColor(BattleMMsg(i).color))
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

'Resize if needed
ResizeScreen

'BLIT THE RESIZABLE BOX BECAUSE WE'RE TOO LAZY TO RECODE THE MAIN ONE

' Scrolling bug here fixed. Man I'm awesome. -Pickle
' SOURCE RECT
rec.Top = 0 + sx
rec.Bottom = 480 + sx
rec.Left = 0 + sx
rec.Right = 640 + sx

' DEST RECT
rec_pos.Top = 0
rec_pos.Left = 0
rec_pos.Bottom = frmMirage.picUber.Height
rec_pos.Right = frmMirage.picUber.Width

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
For X = 0 To MAX_MAPX
For Y = 0 To MAX_MAPY
For i = 1 To MAX_ATTRIBUTE_NPCS
If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_NPC_SPAWN Then
If MapAttributeNpc(i, X, Y).num > 0 Then
'Call ProcessAttributeNpcMovement(i, x, y)
End If
End If
Next i
Next Y
Next X

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

DoEvents
Loop

frmSendGetData.Visible = True
Call SetStatus("Destroying game data...")

'Shutdown the game
Call GameDestroy

If connectionLost Then
    MsgBox "Connection lost!"
    connectionLost = False
End If

Exit Sub

GameErr:
    ' There was a problem in the game loop
    Call DestroyDirectX
    Call StopBGM
    Call frmMirage.Socket.Close
    Call MsgBox("An unexpected error has occured in the game loop. Character has been saved, client shutting down. Error num: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    End

rec.Top = 0
rec.Bottom = MOUSE_HEIGHT
rec.Left = 0
rec.Bottom = MOUSE_WIDTH
Call DD_BackBuffer.BltFast((MouseMoveX), (MouseMoveY), DD_MouseSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)

End Sub

'Clears the back buffer
Sub ClearBackBuffer()
    Dim rec As RECT
    rec.Top = 0
    rec.Left = 0
    rec.Bottom = frmMirage.picUber.Height
    rec.Right = frmMirage.picUber.Width
    Call DD_BackBuffer.BltColorFill(rec, 0)
End Sub

'Draws all the low tiles on the map
Sub DrawLowerMapTiles()
    Dim X As Integer, Y As Integer
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            Call BltTile(X, Y)
        Next Y
    Next X
End Sub

'Draws all the high tiles on the map
Sub DrawUpperMapTiles()
    Dim X As Integer, Y As Integer
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            Call BltFringeTile(X, Y)
            If Not IsTileRoof(X, Y) Then
                Call BltFringe2Tile(X, Y)
            End If
        Next Y
    Next X
End Sub

'TAKEN OUT as it seemed to cause problems. Kept for posterity, but use at your own risk.
Sub BltInventory()
'    Visual Inventory
    Dim Q As Long
    Dim Qq As Long
    Dim IT As Long
    Dim val As Integer
    
'    Set the font
    SetFont "Tahoma", 11
    
    If GetTickCount > IT + 500 And frmMirage.picInv3.Visible = True Then
        For Q = 0 To MAX_INV - 1
            Qq = Player(MyIndex).Inv(Q + 1).num
            
            If frmMirage.picInv(Q).Picture <> LoadPicture() Then
                frmMirage.picInv(Q).Picture = LoadPicture()
            Else
                If Qq = 0 Then
                    frmMirage.picInv(Q).Picture = LoadPicture()
                Else
                    'Draw the item
                    Call BitBlt(frmMirage.picInv(Q).hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(Qq).Pic - Int(item(Qq).Pic / 6) * 6) * PIC_X, Int(item(Qq).Pic / 6) * PIC_Y, SRCCOPY)
                    
                    'Draw the text for stackable items
                    val = GetPlayerInvItemValue(MyIndex, Q + 1)
                    'Make sure it's stackable!
                    If val > 0 And (item(GetPlayerInvItemNum(MyIndex, Q + 1)).Stackable = 1 Or item(GetPlayerInvItemNum(MyIndex, Q + 1)).Type = ITEM_TYPE_CURRENCY) Then
                        If val > 999 Then
                            'Adjust position for large numbers
                            Call DrawText(frmMirage.picInv(Q).hDC, PIC_X - 26, PIC_Y - 12, Trim$(GetPlayerInvItemValue(MyIndex, Q + 1)), RGB(255, 255, 255))
                        Else
                            Call DrawText(frmMirage.picInv(Q).hDC, PIC_X - 16, PIC_Y - 12, Trim$(GetPlayerInvItemValue(MyIndex, Q + 1)), RGB(255, 255, 255))
                        End If
                    End If
                End If
            End If
        Next Q
    End If
    
    'Reset the font
    SetFont Font, fontsize
End Sub

'Resizes the game screen
Sub ResizeScreen()
    
On Error GoTo ResErr
If Screen_RESIZED = 0 Then
    
    If frmMirage.ScrlResolution.Value = 3 Then
     '   frmMirage.Left = 0
     '   frmMirage.Top = 0
     '   frmMirage.Width = PixelsToTwips(1280, 0)
     '   frmMirage.Height = PixelsToTwips(1024 - 28, 1)
     '   frmMirage.picUber.Top = 10
     '   frmMirage.picUber.Width = 960
     '   frmMirage.picUber.Height = 720
     '   frmMirage.picUber.Left = (((1280 - 960) - 160) / 2) + 160
        Screen_RESIZED = 3
     '
     '   frmMirage.txtMyTextBox.Left = 0
     '   frmMirage.txtMyTextBox.Width = 1280 - 0 - 6
     '   frmMirage.txtMyTextBox.Top = frmMirage.picUber.Height + frmMirage.picUber.Top + 5
     '
     '   frmMirage.txtChat.Width = 1280 - 0 - 6
     '   frmMirage.txtChat.Top = frmMirage.txtMyTextBox.Top + 5 + frmMirage.txtMyTextBox.Height
      '  frmMirage.txtChat.Height = TwipsToPixels(frmMirage.Height, 1) - 576 - 15 - 10 - 185 - 5
     '   frmMirage.txtChat.Left = 0
     '
     '   frmMirage.Picture = LoadPicture(App.Path & "\GUI\1280X1024.gif")
    
    End If

    If frmMirage.ScrlResolution.Value = 2 Then
     '   frmMirage.Left = 0
     '   frmMirage.Top = 0
     '   frmMirage.Width = PixelsToTwips(1024, 0)
    '    frmMirage.Height = PixelsToTwips(768 - 28, 1)
    '    frmMirage.picUber.Top = 10
    '    frmMirage.picUber.Width = 768
    '    frmMirage.picUber.Height = 576
    '    frmMirage.picUber.Left = (((1024 - 768) - 160) / 2) + 160
        Screen_RESIZED = 2
    '
    '    frmMirage.txtMyTextBox.Left = 160
    '    frmMirage.txtMyTextBox.Width = 1024 - 160 - 6
    '    frmMirage.txtMyTextBox.Top = frmMirage.picUber.Height + frmMirage.picUber.Top + 5
    '
    '    frmMirage.txtChat.Width = 1024 - 160 - 6
    '    frmMirage.txtChat.Top = frmMirage.txtMyTextBox.Top + 5 + frmMirage.txtMyTextBox.Height
    '    frmMirage.txtChat.Height = TwipsToPixels(frmMirage.Height, 1) - 576 - 15 - 10 - 43 - 5
    '    frmMirage.txtChat.Left = 160
    '
    '    frmMirage.Picture = LoadPicture(App.Path & "\GUI\1024X768.gif")
        
    End If

    If frmMirage.ScrlResolution.Value = 1 Then
      '  frmMirage.Left = 0
       ' frmMirage.Top = 0
     '   frmMirage.Width = PixelsToTwips(805, 0)
     '   frmMirage.Height = PixelsToTwips(600 - 28, 1)
     '   frmMirage.picUber.Top = 0
     '   frmMirage.picUber.Width = 640
     '   frmMirage.picUber.Height = 480
     '   frmMirage.picUber.Left = 160
        Screen_RESIZED = 1
        
     '   frmMirage.txtChat.Width = frmMirage.picUber.Width - 6
     '   frmMirage.txtChat.Top = 15 + frmMirage.picUber.Height + 5
     '   frmMirage.txtChat.Height = 35
     '   frmMirage.txtChat.Left = 160
        
     '   frmMirage.txtMyTextBox.Left = 160
     '   frmMirage.txtMyTextBox.Width = frmMirage.picUber.Width
     '   frmMirage.txtMyTextBox.Top = frmMirage.picUber.Height
        
      '  frmMirage.Picture = LoadPicture(App.Path & "\GUI\800X600.gif")
      '  frmMirage.piccharstats.Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '  frmMirage.picEquip.Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '  frmMirage.picPlayerSpells.Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '  frmMirage.picInv3.Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '  frmMirage.picGuildAdmin.Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '  frmMirage.picWhosOnline.Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '  frmMirage.Picture1(0).Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '  frmMirage.Picture8.Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '  frmMirage.Picture9.Picture = LoadPicture(App.Path & "\GUI\minimenus.gif")
      '
    End If
    End If
    
    Exit Sub
    
' Resolution Change Error Handler
ResErr:
    Call MsgBox("Error resizing screen.", vbInformation)
    frmMirage.ScrlResolution.Value = 1
    Err.Clear
    
End Sub

Sub GameDestroy()
    Call DestroyDirectX
    Call StopBGM
    Unload frmadmin
    End
End Sub

Sub BltTile(ByVal X As Long, ByVal Y As Long)
    
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
    
    Ground = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Ground
    Anim1 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask
    Anim2 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Anim
    Mask2 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask2
    M2Anim = Map(GetPlayerMap(MyIndex)).Tile(X, Y).M2Anim
    
    GroundTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).GroundSet
    MaskTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).MaskSet
    AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).AnimSet
    Mask2TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask2Set
    M2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).M2AnimSet

    ' Only used if ever want to switch to blt rather then bltfast

    With rec_pos
        .Top = (Y - NewPlayerY) * PIC_Y + sx - NewYOffset
        .Bottom = .Top + PIC_Y
        .Left = (X - NewPlayerX) * PIC_X + sx - NewXOffset
        .Right = .Left + PIC_X
    End With
    
    If Int(GroundTileSet) > 10 Then Exit Sub
    If Int(MaskTileSet) > 10 Then Exit Sub
    If Int(Mask2TileSet) > 10 Then Exit Sub
    If Int(AnimTileSet) > 10 Then Exit Sub
    If Int(M2AnimTileSet) > 10 Then Exit Sub
    
    If TileFile(GroundTileSet) = 0 Then Exit Sub
    rec.Top = Int(Ground / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - Int(Ground / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(GroundTileSet), rec, DDBLT_WAIT)
    Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(GroundTileSet), rec, DDBLTFAST_WAIT)
    
    If (MapAnim = 0) Or (Anim2 <= 0) Then
        ' Is there an animation tile to plot?
        If TileFile(MaskTileSet) = 0 Then Exit Sub
        
        If Anim1 > 0 And TempTile(X, Y).DoorOpen = NO Then
            rec.Top = Int(Anim1 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim1 - Int(Anim1 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(MaskTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(MaskTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If

     Else
        ' Is there a second animation tile to plot?

        If Anim2 > 0 Then
            If TileFile(AnimTileSet) = 0 Then Exit Sub
            
            If 0 + ReadINI("CONFIG", "Multiple", App.Path & "\config.ini") = 1 Then
                If Anim2Data >= 7 Then
                    If Anim2Data >= 14 Then
                        If Anim2Data >= 21 Then
                            If Anim2Data >= 28 Then
                                rec.Top = Int(Anim2 / TilesInSheets) * PIC_Y
                                rec.Bottom = rec.Top + PIC_Y
                                rec.Left = val((Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets)) * PIC_X
                                rec.Right = rec.Left + PIC_X
                                
                                'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                                Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                Anim2Data = 0
                                Exit Sub
                             Else
                                rec.Top = Int(Anim2 / TilesInSheets) * PIC_Y
                                rec.Bottom = rec.Top + PIC_Y
                                rec.Left = val((Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets) + 3) * PIC_X
                                rec.Right = rec.Left + PIC_X
                                'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                                Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                Anim2Data = Anim2Data + 1
                                Exit Sub
                            End If

                         Else
                            rec.Top = Int(Anim2 / TilesInSheets) * PIC_Y
                            rec.Bottom = rec.Top + PIC_Y
                            rec.Left = val((Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets) + 2) * PIC_X
                            rec.Right = rec.Left + PIC_X
                            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            Anim2Data = Anim2Data + 1
                            Exit Sub
                        End If

                     Else
                        rec.Top = Int(Anim2 / TilesInSheets) * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        rec.Left = val((Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets) + 1) * PIC_X
                        rec.Right = rec.Left + PIC_X
                        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                        Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        Anim2Data = Anim2Data + 1
                        Exit Sub
                    End If

                 Else
                    rec.Top = Int(Anim2 / TilesInSheets) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    rec.Left = (Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                    Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    Anim2Data = Anim2Data + 1
                    Exit Sub
                End If

             Else
                rec.Top = Int(Anim2 / TilesInSheets) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = (Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets) * PIC_X
                rec.Right = rec.Left + PIC_X
                'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            
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
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(Mask2TileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Mask2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If

     Else
        ' Is there a second animation tile to plot?

        If M2Anim > 0 Then
            If TileFile(M2AnimTileSet) = 0 Then Exit Sub
            
            If 0 + ReadINI("CONFIG", "Multiple", App.Path & "\config.ini") = 1 Then
                If M2AnimData >= 7 Then
                    If M2AnimData >= 14 Then
                        If M2AnimData >= 21 Then
                            If M2AnimData >= 28 Then
                                rec.Top = Int(M2Anim / TilesInSheets) * PIC_Y
                                rec.Bottom = rec.Top + PIC_Y
                                rec.Left = val((M2Anim - Int(M2Anim / TilesInSheets) * TilesInSheets)) * PIC_X
                                rec.Right = rec.Left + PIC_X
                                'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                                Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(M2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                M2AnimData = 0
                                Exit Sub
                            Else
                                rec.Top = Int(M2Anim / TilesInSheets) * PIC_Y
                                rec.Bottom = rec.Top + PIC_Y
                                rec.Left = val((M2Anim - Int(M2Anim / TilesInSheets) * TilesInSheets) + 3) * PIC_X
                                rec.Right = rec.Left + PIC_X
                                'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                                Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(M2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                                M2AnimData = M2AnimData + 1
                                Exit Sub
                            End If
                            
                        Else
                            rec.Top = Int(M2Anim / TilesInSheets) * PIC_Y
                            rec.Bottom = rec.Top + PIC_Y
                            rec.Left = val((M2Anim - Int(M2Anim / TilesInSheets) * TilesInSheets) + 2) * PIC_X
                            rec.Right = rec.Left + PIC_X
                            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(M2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                            M2AnimData = M2AnimData + 1
                            Exit Sub
                        End If
                        
                    Else
                        rec.Top = Int(M2Anim / TilesInSheets) * PIC_Y
                        rec.Bottom = rec.Top + PIC_Y
                        rec.Left = val((M2Anim - Int(M2Anim / TilesInSheets) * TilesInSheets) + 1) * PIC_X
                        rec.Right = rec.Left + PIC_X
                        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                        Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(M2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                        M2AnimData = M2AnimData + 1
                        Exit Sub
                    End If
                    
                Else
                    rec.Top = Int(M2Anim / TilesInSheets) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    rec.Left = (M2Anim - Int(M2Anim / TilesInSheets) * TilesInSheets) * PIC_X
                    rec.Right = rec.Left + PIC_X
                    'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                    Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(M2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    M2AnimData = M2AnimData + 1
                    Exit Sub
                End If
                
            Else
                rec.Top = Int(M2Anim / TilesInSheets) * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                rec.Left = (M2Anim - Int(M2Anim / TilesInSheets) * TilesInSheets) * PIC_X
                rec.Right = rec.Left + PIC_X
                'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(M2AnimTileSet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
                Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(M2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            
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
    
    rec.Top = Int(item(MapItem(ItemNum).num).Pic / 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (item(MapItem(ItemNum).num).Pic - Int(item(MapItem(ItemNum).num).Pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast((MapItem(ItemNum).X - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal X As Long, ByVal Y As Long)
    Dim Fringe As Long
    Dim FAnim As Long
    Dim FringeTileSet As Byte
    Dim FAnimTileSet As Byte
'    Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = Y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = X * PIC_X
        .Right = .Left + PIC_X
    End With
    Fringe = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe
    FAnim = Map(GetPlayerMap(MyIndex)).Tile(X, Y).FAnim
    
    FringeTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).FringeSet
    FAnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).FAnimSet
    
    If Int(FringeTileSet) > ExtraSheets Then Exit Sub
    If Int(FAnimTileSet) > ExtraSheets Then Exit Sub
    
    
    If (MapAnim = 0) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe > 0 Then
            If TileFile(FringeTileSet) = 0 Then Exit Sub
            rec.Top = Int(Fringe / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe - Int(Fringe / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
'            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FringeTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    Else
        
        If FAnim > 0 Then
            If TileFile(FAnimTileSet) = 0 Then Exit Sub
            rec.Top = Int(FAnim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (FAnim - Int(FAnim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X
'            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FAnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If
    
End Sub

Sub BltFringe2Tile(ByVal X As Integer, ByVal Y As Integer)
    Dim Fringe2 As Long
    Dim F2Anim As Long
    Dim Fringe2TileSet As Byte
    Dim F2AnimTileSet As Byte
'    Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = Y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = X * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Fringe2 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe2
    F2Anim = Map(GetPlayerMap(MyIndex)).Tile(X, Y).F2Anim
    
    Fringe2TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe2Set
    F2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).F2AnimSet
    
    If Int(Fringe2TileSet) > ExtraSheets Then Exit Sub
    If Int(F2AnimTileSet) > ExtraSheets Then Exit Sub
    
    If (MapAnim = 0) Or (F2Anim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe2 > 0 Then
        If TileFile(Fringe2TileSet) = 0 Then Exit Sub
        rec.Top = Int(Fringe2 / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (Fringe2 - Int(Fringe2 / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 Then
        If TileFile(F2AnimTileSet) = 0 Then Exit Sub
        rec.Top = Int(F2Anim / TilesInSheets) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (F2Anim - Int(F2Anim / TilesInSheets) * TilesInSheets) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(F2AnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
End Sub

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim X As Long, Y As Integer
Dim AttackSpeed As Long
Dim temp As Long
Dim attack_weaponslot As Long
Dim attack_item As Long

    attack_weaponslot = Int(GetPlayerWeaponSlot(Index))
    
    If attack_weaponslot > 0 Then
    attack_item = Int(Player(Index).Inv(attack_weaponslot).num)
        If attack_item > 0 Then
        AttackSpeed = 1000 'item(attack_item).AttackSpeed
        Else
        AttackSpeed = 1000
        End If
    Else
        AttackSpeed = 1000
    End If

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
        .Bottom = .Top + PIC_Y
        .Left = GetPlayerX(Index) * PIC_X + Player(Index).xOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).yOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).yOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).xOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).xOffset < PIC_Y / 2 * -1) Then Anim = 1
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
    
    temp = GetPlayerShieldSlot(Index)
    If temp > 0 Then
        Player(Index).Shield = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Shield = 0
    End If
    
    temp = GetPlayerArmorSlot(Index)
    If temp > 0 Then
        Player(Index).Armor = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Armor = 0
    End If
    
    temp = GetPlayerHelmetSlot(Index)
    If temp > 0 Then
        Player(Index).Helmet = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Helmet = 0
    End If
    
    temp = GetPlayerWeaponSlot(Index)
    If temp > 0 Then
        Player(Index).Weapon = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Weapon = 0
    End If
    
    temp = GetPlayerRingSlot(Index)
    If temp > 0 Then
        Player(Index).Ring = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Ring = 0
    End If
    
    temp = GetPlayerNecklaceSlot(Index)
    If temp > 0 Then
        Player(Index).Necklace = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).Necklace = 0
    End If
    
    temp = GetPlayerLegsSlot(Index)
    If temp > 0 Then
        Player(Index).legs = GetPlayerInvItemNum(Index, temp)
    Else
        Player(Index).legs = 0
    End If
    
'32 X 64 LOOP
If Spritesize = 1 Then
        
        '32 X 64
        If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then

        rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
        rec.Right = rec.Left + 32

        If Index = MyIndex Then
        X = NewX + sx
        Y = NewY + sx

            'PLAYER 32 X 64 IF DIR = UP
            If GetPlayerDir(MyIndex) = DIR_UP Then
            
                'PLAYER 32 X 64 BLIT SHIELD IF DIR = UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = item(Player(MyIndex).Shield).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 64 BLIT WEAPON IF DIR = UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = item(Player(MyIndex).Weapon).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 64 BLIT NECKLACE IF DIR = UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = item(Player(MyIndex).Necklace).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
        
                'Why was the shield blitted twice here? -Pickle
            End If
                
                If customplayers = 0 Then
                    'PLAYER 32 X 64 BLIT SPRITE
                    rec.Top = GetPlayerSprite(MyIndex) * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    'PLAYER 32 X 64 BLIT SPRITE
                    rec.Top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 64 BLIT LEGS
                If Player(MyIndex).legs > 0 Then
                    rec.Top = item(Player(MyIndex).legs).Pic * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                    
                'PLAYER 32 X 64 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                    rec.Top = item(Player(MyIndex).Armor).Pic * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 64 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                rec.Top = item(Player(MyIndex).Helmet).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            'PLAYER 32 X 64 DIR <> UP
            If GetPlayerDir(MyIndex) <> DIR_UP Then
        
                'PLAYER 32 X 64 BLIT SHIELD IF DIR <> UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = item(Player(MyIndex).Shield).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 64 BLIT WEAPON IF DIR <> UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = item(Player(MyIndex).Weapon).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 64 BLIT NECKLACE IF DIR <> UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = item(Player(MyIndex).Necklace).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            End If
        
        
        '32 X 64 IF OTHER PLAYER
        Else
    
            X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
            Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset
        
            'IF BLIT IS OFFSCREEN ADJUST THE Y VALUE
           '11111 If y < 0 Then
           '     rec.tOp = rec.tOp + (y * -1)
           '     y = 0
           ' End If
            
            'OTHER 32 X 64 IF DIR = UP
            If GetPlayerDir(Index) = DIR_UP Then
                
                'OTHER 32 X 64 BLIT SHIELD IF DIR = UP
                If Player(Index).Shield > 0 Then
                rec.Top = item(Player(Index).Shield).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT WEAPON IF DIR = UP
                If Player(Index).Weapon > 0 Then
                rec.Top = item(Player(Index).Weapon).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT NECKLACE IF DIR = UP
                If Player(Index).Necklace > 0 Then
                rec.Top = item(Player(Index).Necklace).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
   
            End If
            
                'OTHER 32 X 64 BLIT SPRITE
                If 0 + customplayers = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                    
                'OTHER 32 X 64 BLIT LEGS
                If Player(Index).legs > 0 Then
                rec.Top = item(Player(Index).legs).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                rec.Top = item(Player(Index).Armor).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                rec.Top = item(Player(Index).Helmet).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            'OTHER 32 X 64 IF DIR <> UP
            If GetPlayerDir(Index) <> DIR_UP Then
                
                'OTHER 32 X 64 BLIT SHIELD IF DIR <> UP
                If Player(Index).Shield > 0 Then
                rec.Top = item(Player(Index).Shield).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 64 BLIT NECKLACE IF DIR <> UP
                If Player(Index).Necklace > 0 Then
                rec.Top = item(Player(Index).Necklace).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                ''OTHER 32 X 64 BLIT WEAPON IF DIR <> UP
                If Player(Index).Weapon > 0 Then
                rec.Top = item(Player(Index).Weapon).Pic * 64 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
        X = NewX + sx
        Y = NewY + sx
        
        If 0 + customplayers = 0 Then
            'PLAYER 32 X 64 BLIT SPRITE
            rec.Top = GetPlayerSprite(MyIndex) * 64 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            'PLAYER 32 X 64 BLIT SPRITE
            rec.Top = Player(Index).head * 64 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            rec.Top = Player(Index).body * 64 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            rec.Top = Player(Index).leg * 64 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
  
    'OTHER 32 X 64
    Else
        X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
        Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset
        
        'ADJUST IF OFF EDGE OF SCREEN
      '  If y < 0 Then
      '  rec.tOp = rec.tOp + (y * -1)
      '  y = 0
      '11111  End If
        
        'OTHER 32 X 64 BLIT SPRITE
        If 0 + customplayers = 0 Then
            rec.Top = GetPlayerSprite(Index) * 64 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            rec.Top = Player(Index).head * 64 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            rec.Top = Player(Index).body * 64 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            rec.Top = Player(Index).leg * 64 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    'END OF 32 X 64
    End If
   
'32 X 32 LOOP
ElseIf Spritesize = 0 Then

    rec.Top = GetPlayerSprite(Index) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    '32 X 32 PLAYER
    If Index = MyIndex Then
        
        '32 X 32 PAPERDOLLED PLAYER
        If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
            X = NewX + sx
            Y = NewY + sx
        
            'PLAYER 32 X 32 IF DIR = UP
            If GetPlayerDir(MyIndex) = DIR_UP Then
                
                'PLAYER 32 X 32 BLIT SHIELD IF DIR = UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = item(Player(MyIndex).Shield).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 32 X 32 BLIT WEAPON IF DIR = UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = item(Player(MyIndex).Weapon).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT NECKLACE IF DIR = UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = item(Player(MyIndex).Necklace).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            End If
                
                If 0 + customplayers = 0 Then
                    'PLAYER 32 X 32 BLIT SPRITE
                    rec.Top = GetPlayerSprite(MyIndex) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    'PLAYER 32 X 32 BLIT SPRITE
                    rec.Top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT LEGS
                If Player(MyIndex).legs > 0 Then
                rec.Top = item(Player(MyIndex).legs).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                rec.Top = item(Player(MyIndex).Armor).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                rec.Top = item(Player(MyIndex).Helmet).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
            'PLAYER 32 X 32 IF DIR <> UP
            If GetPlayerDir(MyIndex) <> DIR_UP Then
            
                'PLAYER 32 X 32 BLIT SHIELD IF DIR <> UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = item(Player(MyIndex).Shield).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT WEAPON IF DIR <> UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = item(Player(MyIndex).Weapon).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 32 X 32 BLIT NECKLACE IF DIR <> UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = item(Player(MyIndex).Necklace).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            End If
    
            '32 X 32 IF NO PAPERDOLL ON SELF BLIT JUST SPRITE
            Else
            X = NewX + sx
            Y = NewY + sx
                If 0 + customplayers = 0 Then
                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If
    
    '32 X 32 OTHER LOOP
    Else
    X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
    Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset '- 4
            
        'IF OFF TOP EDGE ADJUST
        If Y < 0 Then
        rec.Top = rec.Top + (Y * -1)
        Y = 0
        End If
            
            '32 X 32 OTHER PAPERDOLL LOOP
            If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
            
            'OTHER 32 X 32 IF DIR = UP
            If GetPlayerDir(Index) = DIR_UP Then
                
                'OTHER 32 X 32 BLIT SHIELD IF DIR = UP
                If Player(Index).Shield > 0 Then
                rec.Top = item(Player(Index).Shield).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT WEAPON IF DIR = UP
                If Player(Index).Weapon > 0 Then
                rec.Top = item(Player(Index).Weapon).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT NECKLACE IF DIR = UP
                If Player(Index).Necklace > 0 Then
                rec.Top = item(Player(Index).Necklace).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
            End If
                
                'OTHER 32 X 32 BLIT SPRITE
                If 0 + customplayers = 0 Then
                    rec.Top = GetPlayerSprite(Index) * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                rec.Top = item(Player(Index).Armor).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
        
                'OTHER 32 X 32 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                rec.Top = item(Player(Index).Helmet).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT LEGS
                If Player(Index).legs > 0 Then
                rec.Top = item(Player(Index).legs).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
    
            'OTHER 32 X 32 IF DIR <> UP
            If GetPlayerDir(Index) <> DIR_UP Then
                
                'OTHER 32 X 32 BLIT SHIELD IF DIR <> UP
                If Player(Index).Shield > 0 Then
                rec.Top = item(Player(Index).Shield).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 32 X 32 BLIT WEAPON IF DIR <> UP
                If Player(Index).Weapon > 0 Then
                rec.Top = item(Player(Index).Weapon).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                   
                'OTHER 32 X 32 BLIT NECKLACE IF DIR <> UP
                If Player(Index).Necklace > 0 Then
                rec.Top = item(Player(Index).Necklace).Pic * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
            End If
            
        'OTHER 32 X 32 NON PAPERDOLL
        Else
    
        'OTHER 32 X 32 BLIT NON-PD SPRITE
            If 0 + customplayers = 0 Then
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                rec.Top = Player(Index).head * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                rec.Top = Player(Index).body * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                rec.Top = Player(Index).leg * PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    
End If
Else
        '96 X 96
        If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then

        rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 96
        rec.Right = rec.Left + 96

        If Index = MyIndex Then
        X = val(NewX + sx) - PIC_X
        Y = val(NewY + sx) + PIC_Y

            'PLAYER 96 X 96 IF DIR = UP
            If GetPlayerDir(MyIndex) = DIR_UP Then
            
                'PLAYER 96 X 96 BLIT SHIELD IF DIR = UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = item(Player(MyIndex).Shield).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 96 X 96 BLIT WEAPON IF DIR = UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = item(Player(MyIndex).Weapon).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'PLAYER 96 X 96 BLIT NECKLACE IF DIR = UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = item(Player(MyIndex).Necklace).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
        
                'PLAYER 96 X 96 BLIT SHIELD IF DIR = UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = item(Player(MyIndex).Shield).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If
                
                If customplayers = 0 Then
                    'PLAYER 96 X 96 BLIT SPRITE
                    rec.Top = GetPlayerSprite(MyIndex) * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    'PLAYER 96 X 96 BLIT SPRITE
                    rec.Top = Player(Index).head * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                    
                'PLAYER 96 X 96 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                rec.Top = item(Player(MyIndex).Armor).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 96 X 96 BLIT LEGS
                If Player(MyIndex).legs > 0 Then
                rec.Top = item(Player(MyIndex).legs).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 96 X 96 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                rec.Top = item(Player(MyIndex).Helmet).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            'PLAYER 96 X 96 DIR <> UP
            If GetPlayerDir(MyIndex) <> DIR_UP Then
        
                'PLAYER 96 X 96 BLIT SHIELD IF DIR <> UP
                If Player(MyIndex).Shield > 0 Then
                rec.Top = item(Player(MyIndex).Shield).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 96 X 96 BLIT WEAPON IF DIR <> UP
                If Player(MyIndex).Weapon > 0 Then
                rec.Top = item(Player(MyIndex).Weapon).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
                'PLAYER 96 X 96 BLIT NECKLACE IF DIR <> UP
                If Player(MyIndex).Necklace > 0 Then
                rec.Top = item(Player(MyIndex).Necklace).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            End If
        
        
        '96 X 96 IF OTHER PLAYER
        Else
    
            X = val(GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset) - PIC_X
            Y = val(GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset) + PIC_Y
        
            'IF BLIT IS OFFSCREEN ADJUST THE Y VALUE
           '11111 If y < 0 Then
           '     rec.tOp = rec.tOp + (y * -1)
           '     y = 0
           ' End If
            
            'OTHER 96 X 96 IF DIR = UP
            If GetPlayerDir(Index) = DIR_UP Then
                
                'OTHER 96 X 96 BLIT SHIELD IF DIR = UP
                If Player(Index).Shield > 0 Then
                rec.Top = item(Player(Index).Shield).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 96 X 96 BLIT WEAPON IF DIR = UP
                If Player(Index).Weapon > 0 Then
                rec.Top = item(Player(Index).Weapon).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 96 X 96 BLIT NECKLACE IF DIR = UP
                If Player(Index).Necklace > 0 Then
                rec.Top = item(Player(Index).Necklace).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
   
            End If
            
                'OTHER 96 X 96 BLIT SPRITE
                If 0 + customplayers = 0 Then
                    rec.Top = GetPlayerSprite(Index) * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.Top = Player(Index).head * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).body * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.Top = Player(Index).leg * 96 + PIC_Y
                    rec.Bottom = rec.Top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                    
                'OTHER 96 X 96 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                rec.Top = item(Player(Index).Armor).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 96 X 96 BLIT LEGS
                If Player(Index).legs > 0 Then
                rec.Top = item(Player(Index).legs).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 96 X 96 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                rec.Top = item(Player(Index).Helmet).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            
            'OTHER 96 X 96 IF DIR <> UP
            If GetPlayerDir(Index) <> DIR_UP Then
                
                'OTHER 96 X 96 BLIT SHIELD IF DIR <> UP
                If Player(Index).Shield > 0 Then
                rec.Top = item(Player(Index).Shield).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                'OTHER 96 X 96 BLIT NECKLACE IF DIR <> UP
                If Player(Index).Necklace > 0 Then
                rec.Top = item(Player(Index).Necklace).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                
                ''OTHER 96 X 96 BLIT WEAPON IF DIR <> UP
                If Player(Index).Weapon > 0 Then
                rec.Top = item(Player(Index).Weapon).Pic * 96 + PIC_Y
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
       
            End If
        
        'END OF PAPERDOLL FOR 96 X 96
        End If
        
    'IF 96 X 96 AND NO PAPERDOLL
    Else
        rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 96
        rec.Right = rec.Left + 96
    
    'PLAYER 96 X 96
    If Index = MyIndex Then
        X = NewX + sx
        Y = NewY + sx
        
        If 0 + customplayers = 0 Then
            'PLAYER 96 X 96 BLIT SPRITE
            rec.Top = GetPlayerSprite(MyIndex) * 96 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            'PLAYER 96 X 96 BLIT SPRITE
            rec.Top = Player(Index).head * 96 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            rec.Top = Player(Index).body * 96 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            rec.Top = Player(Index).leg * 96 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
  
    'OTHER 96 X 96
    Else
        X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
        Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset
        
        'ADJUST IF OFF EDGE OF SCREEN
        'If y < 0 Then
        'rec.tOp = rec.tOp + (y * -1)
        'y = 0
        'End If
        
        'OTHER 96 X 96 BLIT SPRITE
        If 0 + customplayers = 0 Then
            rec.Top = GetPlayerSprite(Index) * 96 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            rec.Top = Player(Index).head * 96 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            rec.Top = Player(Index).body * 96 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            rec.Top = Player(Index).leg * 96 + PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    'END OF 96 X 96
    End If
End If

End Sub
Sub BltPlayerTop(ByVal Index As Long)
Dim Anim As Byte
Dim X As Long, Y As Long
Dim AttackSpeed As Long

    If Spritesize = 1 Then
           If GetPlayerWeaponSlot(Index) > 0 Then
               AttackSpeed = 1000 'item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
           Else
               AttackSpeed = 1000
           End If
        
           ' Only used if ever want to switch to blt rather then bltfast
           With rec_pos
               .Top = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
               .Bottom = .Top + PIC_Y
               .Left = GetPlayerX(Index) * PIC_X + Player(Index).xOffset
               .Right = .Left + PIC_X
           End With
           
           ' Check for animation
           Anim = 0
           If Player(Index).Attacking = 0 Then
               Select Case GetPlayerDir(Index)
                   Case DIR_UP
                       If (Player(Index).yOffset < PIC_Y / 2) Then Anim = 1
                   Case DIR_DOWN
                       If (Player(Index).yOffset < PIC_Y / 2 * -1) Then Anim = 1
                   Case DIR_LEFT
                       If (Player(Index).xOffset < PIC_Y / 2) Then Anim = 1
                   Case DIR_RIGHT
                       If (Player(Index).xOffset < PIC_Y / 2 * -1) Then Anim = 1
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
           
        If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
           rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
           rec.Right = rec.Left + PIC_X
        
           If Index = MyIndex Then
           X = NewX + sx
           Y = NewY + sx - 32
           
           
           If GetPlayerDir(Index) = DIR_UP Then
               If Player(MyIndex).Shield > 0 Then
                   rec.Top = item(Player(MyIndex).Shield).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
               If Player(MyIndex).Weapon > 0 Then
                   rec.Top = item(Player(MyIndex).Weapon).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
               If Player(MyIndex).Necklace > 0 Then
                   rec.Top = item(Player(MyIndex).Necklace).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
           End If
               
        
            If 0 + customplayers = 0 Then
                rec.Top = GetPlayerSprite(Index) * 64
                rec.Bottom = rec.Top + PIC_Y
                
                 'If y < 0 Then
                  '      rec.tOp = rec.tOp + (y * -1)
                  '      y = 0
                 'End If
                
                Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                rec.Top = GetPlayerHead(Index) * 64
                rec.Bottom = rec.Top + PIC_Y
                
                ' If y < 0 Then
                '        rec.tOp = rec.tOp + (y * -1)
                '   y = 0
                'End If
                
                Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
            
           If Player(MyIndex).Armor > 0 Then
               rec.Top = item(Player(MyIndex).Armor).Pic * 64
               rec.Bottom = rec.Top + PIC_Y
               Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
           End If
           
           If Player(MyIndex).legs > 0 Then
               rec.Top = item(Player(MyIndex).legs).Pic * 64
               rec.Bottom = rec.Top + PIC_Y
               Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
           End If
        
        If Player(MyIndex).Helmet > 0 Then
               rec.Top = item(Player(MyIndex).Helmet).Pic * 64
               rec.Bottom = rec.Top + PIC_Y
               Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
           End If
           If GetPlayerDir(Index) <> DIR_UP Then
               If Player(MyIndex).Shield > 0 Then
                   rec.Top = item(Player(MyIndex).Shield).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
               If Player(MyIndex).Necklace > 0 Then
                   rec.Top = item(Player(MyIndex).Necklace).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
               If Player(MyIndex).Weapon > 0 Then
                   rec.Top = item(Player(MyIndex).Weapon).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
           End If
        
           
        Else
           X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
           Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32
           
           ' If y < 0 Then
           '        rec.tOp = rec.tOp + (y * -1)
           '        y = 0
           ' End If
           
           If GetPlayerDir(Index) = DIR_UP Then
               If Player(Index).Shield > 0 Then
                   rec.Top = item(Player(Index).Shield).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
               If Player(Index).Necklace > 0 Then
                   rec.Top = item(Player(Index).Necklace).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
               If Player(Index).Weapon > 0 Then
                   rec.Top = item(Player(Index).Weapon).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
           End If
           
           If 0 + customplayers = 0 Then
                rec.Top = GetPlayerSprite(Index) * 64
                rec.Bottom = rec.Top + PIC_Y
                
               '  If y < 0 Then
               '         rec.tOp = rec.tOp + (y * -1)
               '         y = 0
               '  End If
                
                Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                rec.Top = GetPlayerHead(Index) * 64
                rec.Bottom = rec.Top + PIC_Y
                
               '  If y < 0 Then
               '         rec.tOp = rec.tOp + (y * -1)
               '    y = 0
               ' End If
                
                Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
           
           If Player(Index).Armor > 0 Then
               rec.Top = item(Player(Index).Armor).Pic * 64
               rec.Bottom = rec.Top + PIC_Y
               Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
           End If
        
           If Player(Index).legs > 0 Then
               rec.Top = item(Player(Index).legs).Pic * 64
               rec.Bottom = rec.Top + PIC_Y
               Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
           End If
        
           If Player(Index).Helmet > 0 Then
               rec.Top = item(Player(Index).Helmet).Pic * 64
               rec.Bottom = rec.Top + PIC_Y
               Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
           End If
           If GetPlayerDir(Index) <> DIR_UP Then
               If Player(Index).Shield > 0 Then
                   rec.Top = item(Player(Index).Shield).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
               If Player(Index).Necklace > 0 Then
                   rec.Top = item(Player(Index).Necklace).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
               If Player(Index).Weapon > 0 Then
                   rec.Top = item(Player(Index).Weapon).Pic * 64
                   rec.Bottom = rec.Top + PIC_Y
                   Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
           End If
        End If
        Else
           rec.Top = GetPlayerSprite(Index) * 64
           rec.Bottom = rec.Top + PIC_Y
           rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
           rec.Right = rec.Left + PIC_X
        
        If Index = MyIndex Then
           X = NewX + sx
           Y = NewY + sx - 32
           
        Else
           X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
           Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32
           
           ' If y < 0 Then
           '        rec.tOp = rec.tOp + (y * -1)
           '        y = 0
           ' End If
        
           
           Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
               End If
            rec.Top = GetPlayerSprite(Index) * 64
           rec.Bottom = rec.Top + PIC_Y
           
            If 0 + customplayers = 0 Then
                Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                rec.Top = Player(Index).head * 64
                rec.Bottom = rec.Top + PIC_Y
                Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Else
    
   If GetPlayerWeaponSlot(Index) > 0 Then
       AttackSpeed = item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
   Else
       AttackSpeed = 1000
   End If

   ' Only used if ever want to switch to blt rather then bltfast
   With rec_pos
       .Top = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset
       .Bottom = .Top + PIC_Y
       .Left = GetPlayerX(Index) * PIC_X + Player(Index).xOffset
       .Right = .Left + 96
   End With
   
   ' Check for animation
   Anim = 0
   If Player(Index).Attacking = 0 Then
       Select Case GetPlayerDir(Index)
           Case DIR_UP
               If (Player(Index).yOffset < PIC_Y / 2) Then Anim = 1
           Case DIR_DOWN
               If (Player(Index).yOffset < PIC_Y / 2 * -1) Then Anim = 1
           Case DIR_LEFT
               If (Player(Index).xOffset < PIC_Y / 2) Then Anim = 1
           Case DIR_RIGHT
               If (Player(Index).xOffset < PIC_Y / 2 * -1) Then Anim = 1
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
   
If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
   rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
   rec.Right = rec.Left + 96

   If Index = MyIndex Then
   X = NewX + sx - PIC_X
   Y = NewY + sx - 32
   
   If GetPlayerDir(Index) = DIR_UP Then
       If Player(MyIndex).Shield > 0 Then
           rec.Top = item(Player(MyIndex).Shield).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(MyIndex).Weapon > 0 Then
           rec.Top = item(Player(MyIndex).Weapon).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(MyIndex).Necklace > 0 Then
           rec.Top = item(Player(MyIndex).Necklace).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
   End If
       

    If 0 + customplayers = 0 Then
        rec.Top = GetPlayerSprite(Index) * 96
        rec.Bottom = rec.Top + PIC_Y
        
         'If y < 0 Then
         '       rec.tOp = rec.tOp + (y * -1)
         '       y = 0
         'End If
        
        Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.Top = GetPlayerHead(Index) * 96
        rec.Bottom = rec.Top + PIC_Y
        
         'If y < 0 Then
         '       rec.tOp = rec.tOp + (y * -1)
         '       y = 0
         'End If
        
        Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
   If Player(MyIndex).Armor > 0 Then
       rec.Top = item(Player(MyIndex).Armor).Pic * 96
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If
   
   If Player(MyIndex).legs > 0 Then
       rec.Top = item(Player(MyIndex).legs).Pic * 96
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If

If Player(MyIndex).Helmet > 0 Then
       rec.Top = item(Player(MyIndex).Helmet).Pic * 96
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If
   If GetPlayerDir(Index) <> DIR_UP Then
       If Player(MyIndex).Shield > 0 Then
           rec.Top = item(Player(MyIndex).Shield).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(MyIndex).Necklace > 0 Then
           rec.Top = item(Player(MyIndex).Necklace).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(MyIndex).Weapon > 0 Then
           rec.Top = item(Player(MyIndex).Weapon).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
   End If

Else
   X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset - PIC_X
   Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32 + PIC_Y
   
    'If y < 0 Then
    '       rec.tOp = rec.tOp + (y * -1)
    '       y = 0
    'End If
   
   If GetPlayerDir(Index) = DIR_UP Then
       If Player(Index).Shield > 0 Then
           rec.Top = item(Player(Index).Shield).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(Index).Necklace > 0 Then
           rec.Top = item(Player(Index).Necklace).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(Index).Weapon > 0 Then
           rec.Top = item(Player(Index).Weapon).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
   End If

    rec.Top = GetPlayerSprite(Index) * 96
    rec.Bottom = rec.Top + PIC_Y
   
    'If y < 0 Then
    '       rec.tOp = rec.tOp + (y * -1)
    '       y = 0
    'End If
   
   Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   
   If Player(Index).Armor > 0 Then
       rec.Top = item(Player(Index).Armor).Pic * 96
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If

   If Player(Index).legs > 0 Then
       rec.Top = item(Player(Index).legs).Pic * 96
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If

   If Player(Index).Helmet > 0 Then
       rec.Top = item(Player(Index).Helmet).Pic * 96
       rec.Bottom = rec.Top + PIC_Y
       Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
   End If
   If GetPlayerDir(Index) <> DIR_UP Then
       If Player(Index).Shield > 0 Then
           rec.Top = item(Player(Index).Shield).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(Index).Necklace > 0 Then
           rec.Top = item(Player(Index).Necklace).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
       If Player(Index).Weapon > 0 Then
           rec.Top = item(Player(Index).Weapon).Pic * 96
           rec.Bottom = rec.Top + PIC_Y
           Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
   End If
End If
Else
   rec.Top = GetPlayerSprite(Index) * 96
   rec.Bottom = rec.Top + PIC_Y
   rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
   rec.Right = rec.Left + 96

If Index = MyIndex Then
   X = NewX + sx + PIC_X
   Y = NewY + sx - 32 - PIC_Y
   
Else
   X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
   Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32
   
    'If y < 0 Then
    '       rec.tOp = rec.tOp + (y * -1)
    '       y = 0
    'End If

   
   Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
       End If
    rec.Top = GetPlayerSprite(Index) * 96
   rec.Bottom = rec.Top + PIC_Y
   
    If 0 + customplayers = 0 Then
        Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.Top = Player(Index).head * 96
        rec.Bottom = rec.Top + PIC_Y
        Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

End If
    End If
End Sub
Sub BltMapNPCName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long

If Npc(MapNpc(Index).num).Big = 0 And Npc(MapNpc(Index).num).Spritesize = 0 Then
    With Npc(MapNpc(Index).num)
    'Draw name
        TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).yOffset - CLng(PIC_Y / 2) - 4
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
Else
    With Npc(MapNpc(Index).num)
    'Draw name
        TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
        TextY = MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).yOffset - CLng(PIC_Y / 2) - 32
        DrawPlayerNameText TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(.Name), vbWhite
    End With
End If
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim X As Long
Dim Y As Long
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
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2 * -1) Then Anim = 1
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
    
        X = MapNpc(MapNpcNum).X * 32 + sx - 16 + MapNpc(MapNpcNum).xOffset
        Y = MapNpc(MapNpcNum).Y * 32 + sx + MapNpc(MapNpcNum).yOffset
   
        If Y < 0 Then
            modify = -Y
            rec.Top = rec.Top + modify
            rec.Bottom = rec.Top + 32
            Y = 0
        End If
   
        If X < 0 Then
            'rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
            'modify = -X
            'rec.Left = rec.Left + modify - 16
            'rec.Right = rec.Left + 48
            'X = 0
            modify = -X
            rec.Left = rec.Left + modify
            rec.Right = rec.Left + 48
            X = 0
        End If
   
        If 32 + X >= (MAX_MAPX * 32) Then
            modify = X - (MAX_MAPX * 32)
            rec.Left = rec.Left + modify + 16
            rec.Right = rec.Left + 32 - modify
        End If

        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
    
         If Npc(MapNpc(MapNpcNum).num).Spritesize = 1 Then
             rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64 + 32
             rec.Bottom = rec.Top + PIC_Y
             rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
             rec.Right = rec.Left + PIC_X
             
             X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
             Y = MapNpc(MapNpcNum).Y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset
             
             ' Check if its out of bounds because of the offset
             
            If Y < 0 Then
                   rec.Top = rec.Top + (Y * -1)
                   Y = 0
            End If
                    
             'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
             Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
             rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
             rec.Bottom = rec.Top + PIC_Y
             rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
             rec.Right = rec.Left + PIC_X
             
             X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
             Y = MapNpc(MapNpcNum).Y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset
                 
             'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
             Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
         End If
    End If
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
'Moles are cool!
'But a Baron can debug
Dim Anim As Byte
Dim X As Long
Dim Y As Long
Dim NPC_number As Long
Dim modify As Long

    'Get the NPC number
    NPC_number = MapNpc(MapNpcNum).num

   ' Make sure that theres an npc there, and if not exit the sub
   If MapNpc(MapNpcNum).num <= 0 Then
       Exit Sub
   End If
   
    If Npc(NPC_number).Big = 0 Then
        If Npc(MapNpc(MapNpcNum).num).Spritesize = 0 Then Exit Sub
    End If
    
 ' Only used if ever want to switch to blt rather then bltfast
   With rec_pos
       .Top = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).yOffset
       .Bottom = .Top + PIC_Y
       .Left = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).xOffset
       .Right = .Left + PIC_X
   End With
   
   ' Check for animation
   Anim = 0
   If MapNpc(MapNpcNum).Attacking = 0 Then
       Select Case MapNpc(MapNpcNum).Dir
           Case DIR_UP
               If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2) Then Anim = 1
           Case DIR_DOWN
               If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2 * -1) Then Anim = 1
           Case DIR_LEFT
               If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2) Then Anim = 1
           Case DIR_RIGHT
               If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2 * -1) Then Anim = 1
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
       
       X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
       Y = MapNpc(MapNpcNum).Y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset - 32
       
       ' Check if its out of bounds because of the offset
       If Y < 0 Then
           rec.Top = rec.Top + (Y * -1)
           Y = 0
       End If
           
       'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
       Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
 Else
   rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
       
    rec.Top = Npc(MapNpc(MapNpcNum).num).Sprite * 64
    rec.Bottom = rec.Top + 32
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
    rec.Right = rec.Left + 64

    X = MapNpc(MapNpcNum).X * 32 + sx - 16 + MapNpc(MapNpcNum).xOffset
    Y = MapNpc(MapNpcNum).Y * 32 + sx - 32 + MapNpc(MapNpcNum).yOffset

    If Y < 0 Then
        modify = -Y
        rec.Top = rec.Top + modify
        rec.Bottom = rec.Top + 32
        Y = 0
    End If
   
    If X < 0 Then
        'rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
        'modify = -X
        'rec.Left = rec.Left + modify - 16
        'rec.Right = rec.Left + 48
        'X = 0
        modify = -X
        rec.Left = rec.Left + modify
        rec.Right = rec.Left + 48
        X = 0
    End If
   
    If 32 + X >= (MAX_MAPX * 32) Then
        modify = X - (MAX_MAPX * 32)
        rec.Left = rec.Left + modify + 16
        rec.Right = rec.Left + 32 - modify
    End If

    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
End Sub
Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long

    If Player(Index).color <> 0 Then
        If Player(Index).color > 16 Then
            Exit Sub
        Else
            color = QBColor(val(Player(Index).color - 1))
        End If
    Else
        ' Check access level
        If GetPlayerPK(Index) = NO Then
            color = QBColor(Yellow)
            Select Case GetPlayerAccess(Index)
                Case 0
                    color = QBColor(Brown)
                Case 1
                    color = QBColor(DarkGrey)
                Case 2
                    color = QBColor(Cyan)
                Case 3
                    color = QBColor(Blue)
                Case 4
                    color = QBColor(Pink)
            End Select
        Else
            color = QBColor(BrightRed)
        End If
    End If
        
If Spritesize = 1 Then
  If Index = MyIndex Then
        If lvl = 1 Then
            TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
        Else
            TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
        End If
        
        TextY = NewY + sx - 50
        If lvl = 1 Then
           Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index), color)
        Else
           Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), color)
        End If
    Else
       ' Draw name
        If lvl = 1 Then
            TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
        Else
            TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
        End If
        
        TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 32
        
        If lvl = 1 Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index), color)
        Else
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), color)
        End If
    End If
Else
    If Spritesize = 2 Then
        If Index = MyIndex Then
              If lvl = 1 Then
                  TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
              Else
                  TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
              End If
              
              TextY = NewY + sx - 50
              If lvl = 1 Then
                 Call DrawText(TexthDC, TextX, TextY - PIC_Y, GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index), color)
              Else
                 Call DrawText(TexthDC, TextX, TextY - PIC_Y, GetPlayerName(MyIndex), color)
              End If
        Else
             ' Draw name
              If lvl = 1 Then
                  TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
              Else
                  TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
              End If
              
              TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 32
              
              If lvl = 1 Then
                  Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index), color)
              Else
                  Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset - PIC_Y, GetPlayerName(Index), color)
              End If
        End If
    Else
        If Index = MyIndex Then
            If lvl = 1 Then
                TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
            Else
                TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
            End If
            TextY = NewY + sx - Int(PIC_Y / 2)
            
            If lvl = 1 Then
                Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index), color)
            Else
                Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), color)
            End If
        Else
            ' Draw name
            If lvl = 1 Then
                TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
            Else
                TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
            End If
            
            TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2)
            
            If lvl = 1 Then
                Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index), color)
            Else
                Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), color)
            End If
        End If
    End If
End If
End Sub

Sub BltPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long

    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0
                If GetPlayerSTR(Index) > 0 Then
                    color = QBColor(Red)
                Else
                    color = QBColor(Red)
                End If
            Case 1
                color = QBColor(BrightCyan)
            Case 2
                color = QBColor(Yellow)
            Case 3
                color = QBColor(BrightGreen)
            Case 4
                color = QBColor(Yellow)
        End Select
    Else
        color = QBColor(BrightRed)
    End If

If Index = MyIndex Then
TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerGuild(MyIndex)) / 2) * 8)
If Spritesize = 1 Then
TextY = NewY + sx - Int(PIC_Y / 4) - 52
Else
TextY = NewY + sx - Int(PIC_Y / 4) - 20
End If
Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), color)
Else
' Draw name
TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerGuild(Index)) / 2) * 8)
If Spritesize = 1 Then
TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 44
Else
TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 12
End If
Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerGuild(Index), color)
End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
    ' Check if player is walking, and if so process moving them over
    If Player(Index).Moving = MOVING_WALKING Then
            If Player(Index).Access > 0 Then
                If SS_WALK_SPEED <> 0 Then
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            Player(Index).yOffset = Player(Index).yOffset - SS_WALK_SPEED
                        Case DIR_DOWN
                            Player(Index).yOffset = Player(Index).yOffset + SS_WALK_SPEED
                        Case DIR_LEFT
                            Player(Index).xOffset = Player(Index).xOffset - SS_WALK_SPEED
                        Case DIR_RIGHT
                            Player(Index).xOffset = Player(Index).xOffset + SS_WALK_SPEED
                    End Select
                Else
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            Player(Index).yOffset = Player(Index).yOffset - GM_WALK_SPEED
                        Case DIR_DOWN
                            Player(Index).yOffset = Player(Index).yOffset + GM_WALK_SPEED
                        Case DIR_LEFT
                            Player(Index).xOffset = Player(Index).xOffset - GM_WALK_SPEED
                        Case DIR_RIGHT
                            Player(Index).xOffset = Player(Index).xOffset + GM_WALK_SPEED
                    End Select
                End If
            Else
                If SS_WALK_SPEED <> 0 Then
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            Player(Index).yOffset = Player(Index).yOffset - SS_WALK_SPEED
                        Case DIR_DOWN
                            Player(Index).yOffset = Player(Index).yOffset + SS_WALK_SPEED
                        Case DIR_LEFT
                            Player(Index).xOffset = Player(Index).xOffset - SS_WALK_SPEED
                        Case DIR_RIGHT
                            Player(Index).xOffset = Player(Index).xOffset + SS_WALK_SPEED
                    End Select
                Else
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            Player(Index).yOffset = Player(Index).yOffset - WALK_SPEED
                        Case DIR_DOWN
                            Player(Index).yOffset = Player(Index).yOffset + WALK_SPEED
                        Case DIR_LEFT
                            Player(Index).xOffset = Player(Index).xOffset - WALK_SPEED
                        Case DIR_RIGHT
                            Player(Index).xOffset = Player(Index).xOffset + WALK_SPEED
                    End Select
                End If
            End If
            
            ' Check if completed walking over to the next tile
            If (Player(Index).xOffset = 0) And (Player(Index).yOffset = 0) Then
                Player(Index).Moving = 0
            End If
    Else
        ' Check if player is running, and if so process moving them over
        If Player(Index).Moving = MOVING_RUNNING Then
            If GetPlayerSP(Index) > 0 Then
                    Call SetPlayerSP(Index, GetPlayerSP(Index) - 1)
                    frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
                    If Player(Index).Access > 0 Then
                        If SS_RUN_SPEED <> 0 Then
                            Select Case GetPlayerDir(Index)
                                Case DIR_UP
                                    Player(Index).yOffset = Player(Index).yOffset - SS_RUN_SPEED
                                Case DIR_DOWN
                                    Player(Index).yOffset = Player(Index).yOffset + SS_RUN_SPEED
                                Case DIR_LEFT
                                    Player(Index).xOffset = Player(Index).xOffset - SS_RUN_SPEED
                                Case DIR_RIGHT
                                    Player(Index).xOffset = Player(Index).xOffset + SS_RUN_SPEED
                            End Select
                        Else
                            Select Case GetPlayerDir(Index)
                                Case DIR_UP
                                    Player(Index).yOffset = Player(Index).yOffset - GM_RUN_SPEED
                                Case DIR_DOWN
                                    Player(Index).yOffset = Player(Index).yOffset + GM_RUN_SPEED
                                Case DIR_LEFT
                                    Player(Index).xOffset = Player(Index).xOffset - GM_RUN_SPEED
                                Case DIR_RIGHT
                                    Player(Index).xOffset = Player(Index).xOffset + GM_RUN_SPEED
                            End Select
                        End If
                    Else
                        If SS_RUN_SPEED <> 0 Then
                            Select Case GetPlayerDir(Index)
                                Case DIR_UP
                                    Player(Index).yOffset = Player(Index).yOffset - SS_RUN_SPEED
                                Case DIR_DOWN
                                    Player(Index).yOffset = Player(Index).yOffset + SS_RUN_SPEED
                                Case DIR_LEFT
                                    Player(Index).xOffset = Player(Index).xOffset - SS_RUN_SPEED
                                Case DIR_RIGHT
                                    Player(Index).xOffset = Player(Index).xOffset + SS_RUN_SPEED
                            End Select
                        Else
                            Select Case GetPlayerDir(Index)
                                Case DIR_UP
                                    Player(Index).yOffset = Player(Index).yOffset - RUN_SPEED
                                Case DIR_DOWN
                                    Player(Index).yOffset = Player(Index).yOffset + RUN_SPEED
                                Case DIR_LEFT
                                    Player(Index).xOffset = Player(Index).xOffset - RUN_SPEED
                                Case DIR_RIGHT
                                    Player(Index).xOffset = Player(Index).xOffset + RUN_SPEED
                            End Select
                        End If
                    End If
            Else
                ' Bleh too annoying - Pickle
                'Call AddText("You are to tired to run.", Blue)
                Player(Index).Moving = MOVING_WALKING
            End If
        
                
                ' Check if completed walking over to the next tile
                If (Player(Index).xOffset = 0) And (Player(Index).yOffset = 0) Then
                    Player(Index).Moving = 0
                End If
            End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if npc is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).yOffset = MapNpc(MapNpcNum).yOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).xOffset = MapNpc(MapNpcNum).xOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (MapNpc(MapNpcNum).xOffset = 0) And (MapNpc(MapNpcNum).yOffset = 0) Then
            MapNpc(MapNpcNum).Moving = 0
        End If
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim i As Long

MyText = frmMirage.txtMyTextBox.Text

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
    frmMirage.txtMyTextBox.Text = vbNullString
        If Player(MyIndex).Y - 1 > -1 Then
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                Call AddText("The Sign Reads:", Black)
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), Grey)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), Grey)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> vbNullString Then
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
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Emote message
        If Mid$(MyText, 1, 1) = "-" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Player message
        If Mid$(MyText, 1, 1) = "!" Then
            ChatText = Mid$(MyText, 2, Len(MyText) - 1)
            Name = vbNullString
                    
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
            MyText = vbNullString
            Exit Sub
        End If
            
        ' // Commands //
        ' Verification User
        If LCase(Mid(MyText, 1, 5)) = "/info" Then
            ChatText = Mid$(MyText, 6, Len(MyText) - 5)
            Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Whos Online
        If LCase(Mid(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = vbNullString
            Exit Sub
        End If
                        
        ' Checking fps
        If Mid$(MyText, 1, 4) = "/fps" Then
            If BFPS = False Then
                BFPS = True
            Else
                BFPS = False
            End If
            MyText = vbNullString
            Exit Sub
        End If
                
        ' Show inventory
        If LCase(Mid(MyText, 1, 4)) = "/inv" Then
            frmMirage.picInv3.Visible = True
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Request stats
        If LCase(Mid(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
         
        ' Refresh Player
        If LCase(Mid(MyText, 1, 8)) = "/refresh" Then
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline Chat
        If LCase(Mid(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData("dchat" & SEP_CHAR & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Accept Chat
        If LCase(Mid(MyText, 1, 5)) = "/chat" Then
            Call SendData("achat" & SEP_CHAR & END_CHAR)
            MyText = vbNullString
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
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Accept Trade
        If LCase(Mid(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Decline Trade
        If LCase(Mid(MyText, 1, 8)) = "/decline" Then
            Call SendDeclineTrade
            MyText = vbNullString
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
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Join party
        If LCase(Mid(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            MyText = vbNullString
            Exit Sub
        End If
        
        ' Leave party
        If LCase(Mid(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            MyText = vbNullString
            Exit Sub
        End If
        
                ' House Editor
        If LCase(Mid(MyText, 1, 12)) = "/houseeditor" Then
            Call SendRequestEditHouse
            MyText = vbNullString
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' weather command
            If LCase(Mid(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData("weather" & SEP_CHAR & val#(MyText) & SEP_CHAR & END_CHAR)
                    Else
                        If Trim$(LCase(MyText)) = "none" Then i = 0
                        If Trim$(LCase(MyText)) = "rain" Then i = 1
                        If Trim$(LCase(MyText)) = "snow" Then i = 2
                        If Trim$(LCase(MyText)) = "thunder" Then i = 3
                        Call SendData("weather" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If


            ' Clearing a house owner
            If LCase(Mid(MyText, 1, 11)) = "/clearowner" Then
                Call SendData("clearowner" & SEP_CHAR & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Kicking a player
            If LCase(Mid(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Global Message
            If Mid$(MyText, 1, 1) = "'" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Admin Message
            If Mid$(MyText, 1, 1) = "=" Then
                ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If LCase(Mid(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Map Editor
            If LCase(Mid(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Map report
            If LCase(Mid(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Setting sprite
            If LCase(Mid(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Setting player sprite
            If LCase(Mid(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    i = val#(Mid(MyText, 17, 1))
                
                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(i, val#(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Respawn request
            If Mid$(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = vbNullString
                Exit Sub
            End If
        
            ' MOTD change
            If Mid$(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = Mid$(MyText, 7, Len(MyText) - 6)
                    If Trim$(MyText) <> vbNullString Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Check the ban list
            If Mid$(MyText, 1, 3) = "/banlist" Then
                Call SendBanList
                MyText = vbNullString
                Exit Sub
            End If
            ' Reboot the server
            If LCase(Mid(MyText, 1, 7)) = "/reboot" Then
                Call SendData("reboot" & SEP_CHAR & END_CHAR)
                Call GlobalMsg("An Admin Has Started A Server Reboot Please Log Off")
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Banning a player
            If LCase(Mid(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid$(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = vbNullString
                End If
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid$(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Day/Night
            If LCase(Mid(MyText, 1, 9)) = "/daynight" Then
                Call SendData("daynight" & SEP_CHAR & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid$(MyText, 1, 13) = "/editemoticon" Then
                Call SendRequestEditEmoticon
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid$(MyText, 1, 12) = "/editelement" Then
                Call SendRequestEditElement
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing skills request
            If Mid$(MyText, 1, 12) = "/editskill" Then
                Call SendRequestEditSkill
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing quests request
            If Mid$(MyText, 1, 12) = "/editquest" Then
                Call SendRequestEditQuest
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing arrow request
            If Mid$(MyText, 1, 13) = "/editarrow" Then
                Call SendRequestEditArrow
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing npc request
            If Mid$(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNpc
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Editing shop request
            If Mid$(MyText, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                MyText = vbNullString
                Exit Sub
            End If
        
            ' Editing spell request
            If LCase(Trim$(MyText)) = "/editspell" Then
            'If mid$(MyText, 1, 10) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase(Mid(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = val#(Mid(MyText, 12, 1))
                
                MyText = Mid$(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, i)
                MyText = vbNullString
                Exit Sub
            End If
            
            'Reload Scripts
            If LCase(Trim$(MyText)) = "/reload" Then
                Call SendReloadScripts
                MyText = vbNullString
                Exit Sub
            End If
            
            If LCase(Mid(MyText, 1, 9)) = "/editmain" Then
                Call editmain
                MyText = vbNullString
                Exit Sub

            End If
            
            If LCase(Mid(MyText, 1, 9)) = "/spell" Then
                Call firespell(GetPlayerX(MyIndex), GetPlayerY(MyIndex), GetPlayerX(MyIndex) - 10, GetPlayerY(MyIndex), 5, 5, 1000)
                MyText = vbNullString
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase(Mid(MyText, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Tell them its not a valid command
        If Left$(Trim$(MyText), 1) = "/" Then
            For i = 0 To MAX_EMOTICONS
                If Trim$(Emoticons(i).Command) = Trim$(MyText) And Trim$(Emoticons(i).Command) <> "/" Then
                    Call SendData("checkemoticons" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    MyText = vbNullString
                Exit Sub
                End If
            Next i
            Call SendData("checkcommands" & SEP_CHAR & MyText & SEP_CHAR & END_CHAR)
            MyText = vbNullString
        Exit Sub
        End If
            
        ' Say message
        If Len(Trim$(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        MyText = vbNullString
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
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim$(MyText) = vbNullString Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub CheckAttack()
    Dim AttackSpeed As Long
    If GetPlayerWeaponSlot(MyIndex) > 0 Then
        AttackSpeed = item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AttackSpeed
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
Dim i As Long
Dim X As Long
Dim Y As Long

CanMove = True

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

X = GetPlayerX(MyIndex)
Y = GetPlayerY(MyIndex)

If DirUp Then
  Call SetPlayerDir(MyIndex, DIR_UP)
  Y = Y - 1
ElseIf DirDown Then
  Call SetPlayerDir(MyIndex, DIR_DOWN)
  Y = Y + 1
ElseIf DirLeft Then
  Call SetPlayerDir(MyIndex, DIR_LEFT)
  X = X - 1
Else
  Call SetPlayerDir(MyIndex, DIR_RIGHT)
  X = X + 1
End If

Call SendPlayerDir
    

If Y < 0 Then
  If Map(GetPlayerMap(MyIndex)).Up > 0 Then
    Call SendPlayerRequestNewMap(0)
    GettingMap = True
  End If
  CanMove = False
  Exit Function
ElseIf Y > MAX_MAPY Then
  If Map(GetPlayerMap(MyIndex)).Down > 0 Then
    Call SendPlayerRequestNewMap(0)
    GettingMap = True
  End If
  CanMove = False
  Exit Function
ElseIf X < 0 Then
  If Map(GetPlayerMap(MyIndex)).Left > 0 Then
    Call SendPlayerRequestNewMap(0)
    GettingMap = True
  End If
  CanMove = False
  Exit Function
ElseIf X > MAX_MAPX Then
  If Map(GetPlayerMap(MyIndex)).Right > 0 Then
    Call SendPlayerRequestNewMap(0)
    GettingMap = True
  End If
  CanMove = False
  Exit Function
End If

If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_SIGN Or Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_ROOFBLOCK Or Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_SKILL Then
  CanMove = False
  Exit Function
End If

If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_CBLOCK Then
  If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data1 = Player(MyIndex).Class Then Exit Function
  If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data2 = Player(MyIndex).Class Then Exit Function
  If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data3 = Player(MyIndex).Class Then Exit Function
  CanMove = False
End If

If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_GUILDBLOCK And Map(GetPlayerMap(MyIndex)).Tile(X, Y).String1 <> GetPlayerGuild(MyIndex) Then
  CanMove = False
End If

If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_KEY Or Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_DOOR Then
  If TempTile(X, Y).DoorOpen = NO Then
  CanMove = False
  Exit Function
  End If
End If

If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_WALKTHRU Then
  Exit Function
Else
  For i = 1 To MAX_PLAYERS
    If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
      If (GetPlayerX(i) = X) And (GetPlayerY(i) = Y) Then
        CanMove = False
        Exit Function
      End If
    End If
  Next i
End If

For i = 1 To MAX_MAP_NPCS
  If MapNpc(i).num > 0 Then
    If (MapNpc(i).X = X) And (MapNpc(i).Y = Y) Then
      CanMove = False
      Exit Function
    End If
  End If
Next i

If CanAttributeNPCMove(DIR_UP) = False Then
  CanMove = False
  Exit Function
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
                        Player(MyIndex).yOffset = PIC_Y
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                
                    Case DIR_DOWN
                        Call SendPlayerMove
                        Player(MyIndex).yOffset = PIC_Y * -1
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                
                    Case DIR_LEFT
                        Call SendPlayerMove
                        Player(MyIndex).xOffset = PIC_X
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                
                    Case DIR_RIGHT
                        Call SendPlayerMove
                        Player(MyIndex).xOffset = PIC_X * -1
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                End Select
            
                ' Gotta check :)
                If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
                
                ' Close shop window if open
                If frmNewShop.Visible = True Then Unload frmNewShop
                
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
    
    Dim i As Long
    
    InEditor = True
    
    frmMapEditor.Show vbModeless
    EditorSet = 0
    
    For i = 0 To 10
        If frmMapEditor.Option1(i).Value = True Then
            
            If ENCRYPT_TYPE = "BMP" Then
                frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path & "\GFX\Tiles" & i & ".bmp")
            Else
                Set BMU = New BitmapUtils
                strfilename = App.Path & "/gfx/" & "tiles" & i & "." & Trim$(ENCRYPT_TYPE)
                BMU.LoadByteData (strfilename)
                BMU.DecryptByteData (Trim$(ENCRYPT_PASS))
                BMU.DecompressByteData_ZLib
                frmMapEditor.picBackSelect.Cls
'                frmMapEditor.picBackSelect.Width = BMU.ImageWidth
'                frmMapEditor.picBackSelect.Height = BMU.ImageHeight
                Call BMU.Blt(frmMapEditor.picBackSelect.hDC)
            End If
            
            EditorSet = i
        End If
    Next i
    
    If ENCRYPT_TYPE = "BMP" Then
        frmMapEditor.scrlPicture.Max = Int((frmMapEditor.picBackSelect.Height - frmMapEditor.picBack.Height) / PIC_Y)
        frmMapEditor.picBack.Width = 448
    Else
    frmMapEditor.scrlPicture.Max = Int((BMU.ImageHeight - frmMapEditor.picBack.Height) / PIC_Y)
    frmMapEditor.picBack.Width = 448
    'frmMapEditor.Width = BMU.ImageWidth + frmMapEditor.scrlPicture.Width
    End If
End Sub
Public Sub HouseEditorInit()
    Dim BMU As BitmapUtils
    Dim strfilename As String
    Dim i As Long

    InHouseEditor = True
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
    Dim Red As Integer
    Dim Green As Integer
    Dim Blue As Integer
    
    
    
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
    
    'Read news color
    On Error GoTo NewsError
    
    frmMainMenu.picNews.Caption = "Receiving News..."
    
    Exit Sub
    
NewsError:
    'Error reading colors, so set it to white
    frmMainMenu.picNews.Caption = "Recieving News..."
    
End Sub

Public Sub ParseNews()
    Dim Stuff As String
    Dim Stuff2 As String
    Dim Stuff3 As String
    Dim ThisIsANumber As Long
    Dim Red As Integer
    Dim Blue As Integer
    Dim Grn As Integer
    
'    Parse news.ini
    Stuff2 = vbNullString
        Stuff = ReadINI("DATA", "Desc", App.Path & "\News.ini")
        For ThisIsANumber = 1 To Len(Stuff)
           If Mid$(Stuff, ThisIsANumber, 1) = "*" Then
                Stuff2 = Stuff2 & vbCrLf
           Else
                Stuff2 = Stuff2 & Mid$(Stuff, ThisIsANumber, 1)
           End If
        Next
        Stuff3 = vbNullString
        Stuff = ReadINI("DATA", "News", App.Path & "\News.ini")
        For ThisIsANumber = 1 To Len(Stuff)
           If Mid$(Stuff, ThisIsANumber, 1) = "*" Then
                Stuff3 = Stuff3 & vbCrLf
           Else
                Stuff3 = Stuff3 & Mid$(Stuff, ThisIsANumber, 1)
           End If
        Next
        Dim ret
        ret = Chr(13)
        'Set news text
        frmMainMenu.picNews.Caption = Stuff3 & ret & ret & Stuff2
        
        'Set news font & size
        On Error Resume Next
        
        'Parse news color
        On Error GoTo NewsError
        Red = val(ReadINI("COLOR", "Red", App.Path & "\News.ini"))
        Blue = val(ReadINI("COLOR", "Blue", App.Path & "\News.ini"))
        Grn = val(ReadINI("COLOR", "Green", App.Path & "\News.ini"))
        
        'Make sure they're valid
        If Red < 0 Or Red > 255 Or Blue < 0 Or Blue > 255 Or Grn < 0 Or Grn > 255 Then
            frmMainMenu.picNews.ForeColor = RGB(255, 255, 255)
        Else
            Stuff2 = Stuff2 & Mid$(Stuff, ThisIsANumber, 1)
        End If
        Exit Sub
        
NewsError:
        'Error reading colors, so set it to white
        frmMainMenu.picNews.ForeColor = RGB(255, 255, 255)
End Sub


Public Sub EditorMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long, PicX As Long

    If InEditor Then
        x1 = Int(X / PIC_X)
        y1 = Int(Y / PIC_Y)
        
        If frmMapEditor.MousePointer = 2 Then
            If frmMapEditor.mnuType(1).Checked = True Then
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
            ElseIf frmMapEditor.mnuType(3).Checked = True Then
                EditorTileY = Int(Map(GetPlayerMap(MyIndex)).Tile(x1, y1).light / TilesInSheets)
                EditorTileX = (Map(GetPlayerMap(MyIndex)).Tile(x1, y1).light - Int(Map(GetPlayerMap(MyIndex)).Tile(x1, y1).light / TilesInSheets) * TilesInSheets)
                frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
                frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
                frmMapEditor.shpSelected.Height = PIC_Y
                frmMapEditor.shpSelected.Width = PIC_X
            ElseIf frmMapEditor.mnuType(2).Checked = True Then
                With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If .Type = TILE_TYPE_BLOCKED Then frmMapEditor.optBlocked.Value = True
                    If .Type = TILE_TYPE_WALKTHRU Then frmMapEditor.optWalkThru.Value = True
                    If .Type = TILE_TYPE_WARP Then
                        EditorWarpMap = .Data1
                        EditorWarpX = .Data2
                        EditorWarpY = .Data3
                        frmMapEditor.optWarp.Value = True
                    End If
                    If .Type = TILE_TYPE_HEAL Then frmMapEditor.optHeal.Value = True
                    If .Type = TILE_TYPE_ROOFBLOCK Then
                        frmMapEditor.optRoofBlock.Value = True
                        RoofId = .String1
                    End If
                    If .Type = TILE_TYPE_ROOF Then
                        frmMapEditor.optRoof.Value = True
                        RoofId = .String1
                    End If
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
                    If .Type = TILE_TYPE_SCRIPTED Then
                        ScriptNum = .Data1
                        frmMapEditor.optScripted.Value = True
                    End If
                    If .Type = TILE_TYPE_NPC_SPAWN Then
                        NPCSpawnNum = .Data1
                        frmMapEditor.optNPC.Value = True
                    End If
                     If .Type = TILE_TYPE_HOUSE Then
                        HouseItem = .Data1
                        HousePrice = .Data2
                        frmMapEditor.optHouse.Value = True
                    End If
                     If .Type = TILE_TYPE_GUILDBLOCK Then
                        GuildBlock = .Data1
                        frmMapEditor.optGuildBlock.Value = True
                    End If
                     If .Type = TILE_TYPE_CANON Then
                        CanonItem = .Data1
                        CanonDamage = .Data2
                        CanonDirection = .Data3
                        frmMapEditor.optCanon.Value = True
                    End If
                     If .Type = TILE_TYPE_SKILL Then
                        skill1 = .Data1
                        skill2 = .Data2
                        frmMapEditor.optSkill.Value = True
                    End If
                    If .Type = TILE_TYPE_BANK Then frmMapEditor.optBank.Value = True
                    If .Type = TILE_TYPE_HOOKSHOT Then frmMapEditor.OptGHook.Value = True
                    If .Type = TILE_TYPE_ONCLICK Then
                        ClickScript = .Data1
                        frmMapEditor.optClick.Value = True
                    End If
                    If .Type = TILE_TYPE_LOWER_STAT Then
                        MinusHp = .Data1
                        MinusMp = .Data2
                        MinusSp = .Data3
                        MessageMinus = .String1
                        frmMapEditor.optMinusStat.Value = True
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
                    ElseIf frmMapEditor.mnuType(3).Checked = True Then
                        Map(GetPlayerMap(MyIndex)).Tile(x1, y1).light = EditorTileY * TilesInSheets + EditorTileX
                    ElseIf frmMapEditor.mnuType(2).Checked = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If frmMapEditor.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                            If frmMapEditor.optRoofBlock.Value = True Then
                                .Type = TILE_TYPE_ROOFBLOCK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = RoofId
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optRoof.Value = True Then
                                .Type = TILE_TYPE_ROOF
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = RoofId
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optWarp.Value = True Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
        
                            If frmMapEditor.optHeal.Value = True Then
                                .Type = TILE_TYPE_HEAL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
        
                            If frmMapEditor.optKill.Value = True Then
                                .Type = TILE_TYPE_KILL
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                            If frmMapEditor.optItem.Value = True Then
                                .Type = TILE_TYPE_ITEM
                                .Data1 = ItemEditorNum
                                .Data2 = ItemEditorValue
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            End If
                            If frmMapEditor.optNpcAvoid.Value = True Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKey.Value = True Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optKeyOpen.Value = True Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                                .String1 = KeyOpenEditorMsg
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optShop.Value = True Then
                                .Type = TILE_TYPE_SHOP
                                .Data1 = EditorShopNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optCBlock.Value = True Then
                                .Type = TILE_TYPE_CBLOCK
                                .Data1 = EditorItemNum1
                                .Data2 = EditorItemNum2
                                .Data3 = EditorItemNum3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optArena.Value = True Then
                                .Type = TILE_TYPE_ARENA
                                .Data1 = Arena1
                                .Data2 = Arena2
                                .Data3 = Arena3
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSound.Value = True Then
                                .Type = TILE_TYPE_SOUND
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = SoundFileName
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSprite.Value = True Then
                                .Type = TILE_TYPE_SPRITE_CHANGE
                                .Data1 = SpritePic
                                .Data2 = SpriteItem
                                .Data3 = SpritePrice
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
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
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
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
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If GetPlayerAccess(MyIndex) >= 3 Then
                            If frmMapEditor.optClassChange.Value = True Then
                                .Type = TILE_TYPE_CLASS_CHANGE
                                .Data1 = ClassChange
                                .Data2 = ClassChangeReq
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            End If
                            If frmMapEditor.optScripted.Value = True Then
                                .Type = TILE_TYPE_SCRIPTED
                                .Data1 = ScriptNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optCanon.Value = True Then
                                .Type = TILE_TYPE_CANON
                                .Data1 = CanonItem
                                .Data2 = CanonDamage
                                .Data3 = CanonDirection
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optSkill.Value = True Then
                                .Type = TILE_TYPE_SKILL
                                .Data1 = skill1
                                .Data2 = skill2
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optNPC.Value = True Then
                                .Type = TILE_TYPE_NPC_SPAWN
                                .Data1 = NPCSpawnNum
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optHouse.Value = True Then
                                .Type = TILE_TYPE_HOUSE
                                .Data1 = HouseItem
                                .Data2 = HousePrice
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optGuildBlock.Value = True Then
                                .Type = TILE_TYPE_GUILDBLOCK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = GuildBlock
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optBank.Value = True Then
                                .Type = TILE_TYPE_BANK
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.OptGHook.Value = True Then
                                .Type = TILE_TYPE_HOOKSHOT
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optWalkThru.Value = True Then
                                .Type = TILE_TYPE_WALKTHRU
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optClick.Value = True Then
                                .Type = TILE_TYPE_ONCLICK
                                .Data1 = ClickScript
                                .Data2 = 0
                                .Data3 = 0
                                .String1 = vbNullString
                                .String2 = vbNullString
                                .String3 = vbNullString
                            End If
                            If frmMapEditor.optMinusStat.Value = True Then
                                .Type = TILE_TYPE_LOWER_STAT
                                .Data1 = MinusHp
                                .Data2 = MinusMp
                                .Data3 = MinusSp
                                .String1 = MessageMinus
                                .String2 = vbNullString
                                .String3 = vbNullString
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
                                    ElseIf frmMapEditor.mnuType(3).Checked = True Then
                                        Map(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2).light = (EditorTileY + y2) * TilesInSheets + (EditorTileX + x2)
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
                ElseIf frmMapEditor.mnuType(3).Checked = True Then
                    Map(GetPlayerMap(MyIndex)).Tile(x1, y1).light = 0
                ElseIf frmMapEditor.mnuType(2).Checked = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End With
                End If
            End If
        End If
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        EditorTileX = Int(X / PIC_X)
        EditorTileY = Int(Y / PIC_Y)
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
    frmMirage.Show
    frmMapEditor.MousePointer = 1
    frmMirage.MousePointer = 1
    LoadMap (GetPlayerMap(MyIndex))
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Ground = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).GroundSet = 0
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).MaskSet = 0
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).AnimSet = 0
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask2Set = 0
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).M2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).M2AnimSet = 0
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).FringeSet = 0
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).FAnim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).FAnimSet = 0
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe2 = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe2Set = 0
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
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).F2Anim = 0
                    Map(GetPlayerMap(MyIndex)).Tile(X, Y).F2AnimSet = 0
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
                Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = 0
            Next X
        Next Y
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

Public Sub skillEditorInit()
Dim i As Long
Dim j As Long

    frmskilleditor.txtName.Text = skill(EditorIndex).Name
    frmskilleditor.txtAction.Text = Trim$(skill(EditorIndex).Action)
    frmskilleditor.txtFail.Text = Trim$(skill(EditorIndex).Fail)
    frmskilleditor.txtSucces.Text = Trim$(skill(EditorIndex).Succes)
    frmskilleditor.txtAttempt.Text = Trim$(skill(EditorIndex).AttemptName)
    frmskilleditor.HScroll1.Value = val(skill(EditorIndex).Pictop)
    frmskilleditor.HScroll2.Value = val(skill(EditorIndex).Picleft)
    
    For j = 0 To 4
        frmskilleditor.cmbItem(j).addItem "None"
        For i = 1 To MAX_ITEMS
            frmskilleditor.cmbItem(j).addItem i & ": " & item(i).Name
        Next i
    Next j
    
    For j = 1 To MAX_SKILLS_SHEETS
        frmskilleditor.cmbLevel.addItem "All levels"
        For i = 1 To MAX_SKILL_LEVEL
            frmskilleditor.cmbLevel.addItem "Level " & i
        Next i
            ItemTake1num(j) = skill(EditorIndex).ItemTake1num(j)
            ItemTake2num(j) = skill(EditorIndex).ItemTake2num(j)
            ItemGive1num(j) = skill(EditorIndex).ItemGive1num(j)
            ItemGive2num(j) = skill(EditorIndex).ItemGive2num(j)
            ItemTake1val(j) = skill(EditorIndex).ItemTake1val(j)
            ItemTake2val(j) = skill(EditorIndex).ItemTake2val(j)
            ItemGive1val(j) = skill(EditorIndex).ItemGive1val(j)
            ItemGive2val(j) = skill(EditorIndex).ItemGive2val(j)
            itemequiped(j) = skill(EditorIndex).itemequiped(j)
    Next j
    
    currentsheet = 1
    frmskilleditor.cmbItem(0).ListIndex = itemequiped(currentsheet)
    frmskilleditor.cmbItem(1).ListIndex = ItemTake1num(currentsheet)
    frmskilleditor.cmbItem(2).ListIndex = ItemTake2num(currentsheet)
    frmskilleditor.cmbItem(3).ListIndex = ItemGive1num(currentsheet)
    frmskilleditor.cmbItem(4).ListIndex = ItemGive2num(currentsheet)
    frmskilleditor.cmbLevel.ListIndex = minlevel(currentsheet)
    frmskilleditor.HScroll3.Value = ItemTake1val(currentsheet)
    frmskilleditor.HScroll4.Value = ItemTake2val(currentsheet)
    frmskilleditor.HScroll5.Value = ItemGive1val(currentsheet)
    frmskilleditor.HScroll6.Value = ItemGive2val(currentsheet)
    frmskilleditor.HScroll7.Value = ExpGiven(currentsheet)
    frmskilleditor.HScroll8.Value = base_chance(currentsheet)
    frmskilleditor.Label11.Caption = ItemTake1val(currentsheet)
    frmskilleditor.Label14.Caption = ItemTake2val(currentsheet)
    frmskilleditor.Label17.Caption = ItemGive1val(currentsheet)
    frmskilleditor.Label20.Caption = ItemGive2val(currentsheet)
    frmskilleditor.Label25.Caption = ExpGiven(currentsheet)
    frmskilleditor.Label28.Caption = base_chance(currentsheet)
    
    frmskilleditor.Show vbModal
End Sub

Public Sub SkillEditorOk()
Dim j As Long

    skill(EditorIndex).Name = frmskilleditor.txtName.Text
    
    skill(EditorIndex).Action = frmskilleditor.txtAction.Text
    
    skill(EditorIndex).Fail = frmskilleditor.txtFail.Text
    skill(EditorIndex).Succes = frmskilleditor.txtSucces.Text
    skill(EditorIndex).AttemptName = frmskilleditor.txtAttempt.Text
    
    skill(EditorIndex).Pictop = frmskilleditor.HScroll1.Value
    skill(EditorIndex).Picleft = frmskilleditor.HScroll2.Value

    For j = 1 To MAX_SKILLS_SHEETS
        skill(EditorIndex).ItemTake1num(j) = ItemTake1num(j)
        skill(EditorIndex).ItemTake2num(j) = ItemTake2num(j)
        skill(EditorIndex).ItemGive1num(j) = ItemGive1num(j)
        skill(EditorIndex).ItemGive2num(j) = ItemGive2num(j)
        skill(EditorIndex).minlevel(j) = minlevel(j)
        skill(EditorIndex).ExpGiven(j) = ExpGiven(j)
        skill(EditorIndex).base_chance(j) = base_chance(j)
        skill(EditorIndex).ItemTake1val(j) = ItemTake1val(j)
        skill(EditorIndex).ItemTake2val(j) = ItemTake2val(j)
        skill(EditorIndex).ItemGive1val(j) = ItemGive1val(j)
        skill(EditorIndex).ItemGive2val(j) = ItemGive2val(j)
        skill(EditorIndex).itemequiped(j) = itemequiped(j)
    Next j
    
    Call SendSaveSkill(EditorIndex)
    Call SkillEditorCancel
End Sub

Public Sub SkillEditorCancel()
    InSkillEditor = False
    Unload frmskilleditor
End Sub


Public Sub QuestEditorInit()
Dim i As Long
Dim j As Long

    FrmQuestEditor.txtName.Text = Quest(EditorIndex).Name
    FrmQuestEditor.HScroll1.Value = val(Quest(EditorIndex).Pictop)
    FrmQuestEditor.HScroll2.Value = val(Quest(EditorIndex).Picleft)
    
    For j = 1 To 4
        FrmQuestEditor.cmbItem(j).addItem "None"
        For i = 1 To MAX_ITEMS
            FrmQuestEditor.cmbItem(j).addItem i & ": " & item(i).Name
        Next i
    Next j
    
    For j = 0 To MAX_QUEST_LENGHT
        Q_Map(j) = Quest(EditorIndex).Map(j)
        Q_X(j) = Quest(EditorIndex).X(j)
        Q_Y(j) = Quest(EditorIndex).Y(j)
        Q_Npc(j) = Quest(EditorIndex).Npc(j)
        Q_Script(j) = Quest(EditorIndex).Script(j)
        Q_ItemTake1num(j) = Quest(EditorIndex).ItemTake1num(j)
        Q_ItemTake2num(j) = Quest(EditorIndex).ItemTake2num(j)
        Q_ItemGive1num(j) = Quest(EditorIndex).ItemGive1num(j)
        Q_ItemGive2num(j) = Quest(EditorIndex).ItemGive2num(j)
        Q_ItemTake1val(j) = Quest(EditorIndex).ItemTake1val(j)
        Q_ItemTake2val(j) = Quest(EditorIndex).ItemTake2val(j)
        Q_ItemGive1val(j) = Quest(EditorIndex).ItemGive1val(j)
        Q_ItemGive2val(j) = Quest(EditorIndex).ItemGive2val(j)
        Q_ExpGiven(j) = Quest(EditorIndex).ExpGiven(j)
    Next j

    currentsheet = 0

    FrmQuestEditor.HScroll8.Value = Q_Map(currentsheet)
    FrmQuestEditor.HScroll9.Value = Q_X(currentsheet)
    FrmQuestEditor.HScroll10.Value = Q_Npc(currentsheet)
    FrmQuestEditor.HScroll11.Value = Q_Y(currentsheet)
    FrmQuestEditor.HScroll12.Value = Q_Script(currentsheet)
    FrmQuestEditor.HScroll3.Value = Q_ItemTake1val(currentsheet)
    FrmQuestEditor.HScroll4.Value = Q_ItemTake2val(currentsheet)
    FrmQuestEditor.HScroll5.Value = Q_ItemGive1val(currentsheet)
    FrmQuestEditor.HScroll6.Value = Q_ItemGive2val(currentsheet)
    FrmQuestEditor.HScroll7.Value = Q_ExpGiven(currentsheet)

    FrmQuestEditor.Label11.Caption = FrmQuestEditor.HScroll3.Value
    FrmQuestEditor.Label14.Caption = FrmQuestEditor.HScroll4.Value
    FrmQuestEditor.Label17.Caption = FrmQuestEditor.HScroll5.Value
    FrmQuestEditor.Label2.Caption = FrmQuestEditor.HScroll6.Value
    FrmQuestEditor.Label25.Caption = FrmQuestEditor.HScroll7.Value
    FrmQuestEditor.Label3.Caption = FrmQuestEditor.HScroll8.Value
    FrmQuestEditor.Label20.Caption = FrmQuestEditor.HScroll9.Value
    FrmQuestEditor.Label24.Caption = FrmQuestEditor.HScroll10.Value
    FrmQuestEditor.Label28.Caption = FrmQuestEditor.HScroll11.Value
    FrmQuestEditor.Label31.Caption = FrmQuestEditor.HScroll12.Value

    If Q_ItemTake1num(currentsheet) <> 0 Then
        FrmQuestEditor.cmbItem(1).ListIndex = Q_ItemTake1num(currentsheet)
    Else
        FrmQuestEditor.cmbItem(1).ListIndex = 0
    End If
    If Q_ItemTake2num(currentsheet) <> 0 Then
        FrmQuestEditor.cmbItem(2).ListIndex = Q_ItemTake2num(currentsheet)
    Else
        FrmQuestEditor.cmbItem(2).ListIndex = 0
    End If
    If Q_ItemGive1num(currentsheet) <> 0 Then
        FrmQuestEditor.cmbItem(3).ListIndex = Q_ItemGive1num(currentsheet)
    Else
        FrmQuestEditor.cmbItem(3).ListIndex = 0
    End If
    If Q_ItemGive2num(currentsheet) <> 0 Then
        FrmQuestEditor.cmbItem(4).ListIndex = Q_ItemGive2num(currentsheet)
    Else
        FrmQuestEditor.cmbItem(4).ListIndex = 0
    End If


    FrmQuestEditor.Show vbModal
End Sub

Public Sub QuestEditorOk()
Dim j As Long

    Quest(EditorIndex).Name = FrmQuestEditor.txtName.Text
    Quest(EditorIndex).Pictop = FrmQuestEditor.HScroll1.Value
    Quest(EditorIndex).Picleft = FrmQuestEditor.HScroll2.Value

    For j = 0 To MAX_QUEST_LENGHT
        Quest(EditorIndex).Map(j) = Q_Map(j)
        Quest(EditorIndex).X(j) = Q_X(j)
        Quest(EditorIndex).Y(j) = Q_Y(j)
        Quest(EditorIndex).Npc(j) = Q_Npc(j)
        Quest(EditorIndex).Script(j) = Q_Script(j)
        Quest(EditorIndex).ItemTake1num(j) = Q_ItemTake1num(j)
        Quest(EditorIndex).ItemTake2num(j) = Q_ItemTake2num(j)
        Quest(EditorIndex).ItemGive1num(j) = Q_ItemGive1num(j)
        Quest(EditorIndex).ItemGive2num(j) = Q_ItemGive2num(j)
        Quest(EditorIndex).ItemTake1val(j) = Q_ItemTake1val(j)
        Quest(EditorIndex).ItemTake2val(j) = Q_ItemTake2val(j)
        Quest(EditorIndex).ItemGive1val(j) = Q_ItemGive1val(j)
        Quest(EditorIndex).ItemGive2val(j) = Q_ItemGive2val(j)
        Quest(EditorIndex).ExpGiven(j) = Q_ExpGiven(j)
    Next j
    
    Call SendSaveQuest(EditorIndex)
    Call QuestEditorCancel
End Sub

Public Sub QuestEditorCancel()
    InQuestEditor = False
    Unload FrmQuestEditor
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.Value
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = vbNullString
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
    EditorItemY = Int(item(EditorIndex).Pic / 6)
    EditorItemX = (item(EditorIndex).Pic - Int(item(EditorIndex).Pic / 6) * 6)
    
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
    
    frmItemEditor.txtName.Text = Trim$(item(EditorIndex).Name)
    frmItemEditor.txtDesc.Text = Trim$(item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = item(EditorIndex).Type
    frmItemEditor.txtPrice.Text = item(EditorIndex).Price
    frmItemEditor.chkStackable.Value = item(EditorIndex).Stackable
    frmItemEditor.chkBound.Value = item(EditorIndex).Bound
    
 If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
frmItemEditor.fraEquipment.Visible = True
frmItemEditor.fraAttributes.Visible = True
If frmItemEditor.cmbType.ListIndex = ITEM_TYPE_WEAPON Then
frmItemEditor.fraBow.Visible = True
End If
        
        frmItemEditor.scrlDurability.Value = item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = item(EditorIndex).SpeedReq
        frmItemEditor.scrlClassReq.Value = item(EditorIndex).ClassReq
        frmItemEditor.scrlAccessReq.Value = item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.Value = item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.Value = item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.Value = item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.Value = item(EditorIndex).AddStr
        frmItemEditor.scrlAddDef.Value = item(EditorIndex).AddDef
        frmItemEditor.scrlAddMagi.Value = item(EditorIndex).AddMagi
        frmItemEditor.scrlAddSpeed.Value = item(EditorIndex).AddSpeed
'        frmItemEditor.scrlAddEXP.Value = Item(EditorIndex).AddEXP
        frmItemEditor.scrlAttackSpeed.Value = item(EditorIndex).AttackSpeed
        
        If item(EditorIndex).Data3 > 0 Then
            If item(EditorIndex).Stackable = 1 Then
                frmItemEditor.chkBow.Value = Checked
                frmItemEditor.chkGrapple.Value = Checked
            Else
                frmItemEditor.chkBow.Value = Checked
                frmItemEditor.chkGrapple.Value = Unchecked
            End If
        Else
            frmItemEditor.chkBow.Value = Unchecked
        End If
        
        
        frmItemEditor.cmbBow.Clear
        If frmItemEditor.chkBow.Value = Checked Then
            For i = 1 To 100
                frmItemEditor.cmbBow.addItem i & ": " & Arrows(i).Name
            Next i
            frmItemEditor.cmbBow.ListIndex = item(EditorIndex).Data3 - 1
            frmItemEditor.picBow.Top = (Arrows(item(EditorIndex).Data3).Pic * 32) * -1
            frmItemEditor.cmbBow.Enabled = True
        Else
            frmItemEditor.cmbBow.addItem "None"
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
        frmItemEditor.scrlVitalMod.Value = item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = item(EditorIndex).Data1
        frmItemEditor.chkStackable.Visible = False
    Else
        frmItemEditor.fraSpell.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_SCRIPTED) Then
        frmItemEditor.fraScript.Visible = True
        frmItemEditor.scrlScript.Value = item(EditorIndex).Data1
        frmItemEditor.chkStackable.Visible = True
    Else
        frmItemEditor.fraScript.Visible = False
    End If
    frmItemEditor.VScroll1.Value = EditorItemY
    frmItemEditor.picItems.Top = (EditorItemY) * -32
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    item(EditorIndex).Name = frmItemEditor.txtName.Text
    item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex
    item(EditorIndex).Price = val#(frmItemEditor.txtPrice.Text)
    item(EditorIndex).Bound = frmItemEditor.chkBound.Value
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_NECKLACE) Then
        item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        If frmItemEditor.chkBow.Value = Checked Then
            If frmItemEditor.chkGrapple.Value = Checked Then
                item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                item(EditorIndex).Stackable = 1
            Else
                item(EditorIndex).Data3 = frmItemEditor.cmbBow.ListIndex + 1
                item(EditorIndex).Stackable = 0
            End If
        Else
            item(EditorIndex).Data3 = 0
            item(EditorIndex).Stackable = 0
        End If
        item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        
        item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.Value
        item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.Value
        
        item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.Value
        item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.Value
        item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.Value
        item(EditorIndex).AddStr = frmItemEditor.scrlAddStr.Value
        item(EditorIndex).AddDef = frmItemEditor.scrlAddDef.Value
        item(EditorIndex).AddMagi = frmItemEditor.scrlAddMagi.Value
        item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.Value
        item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.Value
        item(EditorIndex).AttackSpeed = frmItemEditor.scrlAttackSpeed.Value
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
        item(EditorIndex).Data2 = 0
        item(EditorIndex).Data3 = 0
        item(EditorIndex).StrReq = 0
        item(EditorIndex).DefReq = 0
        item(EditorIndex).SpeedReq = 0
        item(EditorIndex).ClassReq = -1
        item(EditorIndex).AccessReq = 0
        
        item(EditorIndex).AddHP = 0
        item(EditorIndex).AddMP = 0
        item(EditorIndex).AddSP = 0
        item(EditorIndex).AddStr = 0
        item(EditorIndex).AddDef = 0
        item(EditorIndex).AddMagi = 0
        item(EditorIndex).AddSpeed = 0
        item(EditorIndex).AddEXP = 0
        item(EditorIndex).AttackSpeed = 0
        item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
        
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_NONE) Then
        item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        item(EditorIndex).Data2 = 0
        item(EditorIndex).Data3 = 0
        item(EditorIndex).StrReq = 0
        item(EditorIndex).DefReq = 0
        item(EditorIndex).SpeedReq = 0
        item(EditorIndex).ClassReq = -1
        item(EditorIndex).AccessReq = 0
        
        item(EditorIndex).AddHP = 0
        item(EditorIndex).AddMP = 0
        item(EditorIndex).AddSP = 0
        item(EditorIndex).AddStr = 0
        item(EditorIndex).AddDef = 0
        item(EditorIndex).AddMagi = 0
        item(EditorIndex).AddSpeed = 0
        item(EditorIndex).AddEXP = 0
        item(EditorIndex).AttackSpeed = 0
        item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCRIPTED) Then
        item(EditorIndex).Data1 = frmItemEditor.scrlScript.Value
        item(EditorIndex).Data2 = 0
        item(EditorIndex).Data3 = 0
        item(EditorIndex).StrReq = 0
        item(EditorIndex).DefReq = 0
        item(EditorIndex).SpeedReq = 0
        item(EditorIndex).ClassReq = -1
        item(EditorIndex).AccessReq = 0
        
        item(EditorIndex).AddHP = 0
        item(EditorIndex).AddMP = 0
        item(EditorIndex).AddSP = 0
        item(EditorIndex).AddStr = 0
        item(EditorIndex).AddDef = 0
        item(EditorIndex).AddMagi = 0
        item(EditorIndex).AddSpeed = 0
        item(EditorIndex).AddEXP = 0
        item(EditorIndex).AttackSpeed = 0
        item(EditorIndex).Stackable = frmItemEditor.chkStackable.Value
    End If
          If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_THROW) Then
        item(EditorIndex).Data1 = frmItemEditor.scrlScript.Value
        item(EditorIndex).Data2 = 0
        item(EditorIndex).Data3 = 0
        item(EditorIndex).StrReq = 0
        item(EditorIndex).DefReq = 0
        item(EditorIndex).SpeedReq = 0
        item(EditorIndex).ClassReq = -1
        item(EditorIndex).AccessReq = 0
        
        item(EditorIndex).AddHP = 0
        item(EditorIndex).AddMP = 0
        item(EditorIndex).AddSP = 0
        item(EditorIndex).AddStr = 0
        item(EditorIndex).AddDef = 0
        item(EditorIndex).AddMagi = 0
        item(EditorIndex).AddSpeed = 0
        item(EditorIndex).AddEXP = 0
        item(EditorIndex).AttackSpeed = 0
        item(EditorIndex).Stackable = 0
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
    frmNpcEditor.scrlSPEED.Value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.Value = Npc(EditorIndex).Big
    frmNpcEditor.StartHP.Value = Npc(EditorIndex).MaxHp
    frmNpcEditor.ExpGive.Value = Npc(EditorIndex).Exp
    frmNpcEditor.txtChance.Text = STR(Npc(EditorIndex).ItemNPC(1).chance)
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
        If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SCRIPTED Then
    frmNpcEditor.scrlScript.Value = Npc(EditorIndex).SpawnSecs
    frmNpcEditor.scrlElement.Value = Npc(EditorIndex).Element
    End If
    If val(0 + Npc(EditorIndex).Spritesize) = 0 Then
        frmNpcEditor.Opt32.Value = 1
        frmNpcEditor.Opt64.Value = 0
    Else
        frmNpcEditor.Opt64.Value = 1
        frmNpcEditor.Opt32.Value = 0
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
Npc(EditorIndex).SpawnSecs = val#(frmNpcEditor.txtSpawnSecs.Text)
Else
Npc(EditorIndex).SpawnSecs = frmNpcEditor.scrlScript.Value
End If
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.Value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.Value
    Npc(EditorIndex).MaxHp = frmNpcEditor.StartHP.Value
    Npc(EditorIndex).Exp = frmNpcEditor.ExpGive.Value
    
    If frmNpcEditor.Opt64.Value = True Then
        Npc(EditorIndex).Spritesize = 1
    Else
        Npc(EditorIndex).Spritesize = 0
    End If
    
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


    If frmNpcEditor.BigNpc.Value = Checked Then

        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * 64

        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, 64, 64, frmNpcEditor.picSprites.hDC, 3 * 64, frmNpcEditor.scrlSprite.Value * 64, SRCCOPY)
    Else
    If Spritesize = 1 Then
    
        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * 64
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, 64, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.Value * 64, SRCCOPY)
    Else
    
        frmNpcEditor.picSprites.Top = frmNpcEditor.scrlSprite.Value * PIC_Y
        
        Call BitBlt(frmNpcEditor.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picSprites.hDC, 3 * PIC_X, frmNpcEditor.scrlSprite.Value * PIC_Y, SRCCOPY)
    End If
    End If
End Sub

' Initializes the shop editor
Public Sub ShopEditorInit()
Dim i As Integer
Dim itemN As Integer
Dim cItemMade As Boolean

   On Error GoTo ShopEditorInit_Error

    
    frmShopEditor.txtName.Text = Trim$(Shop(EditorIndex).Name)
    frmShopEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    frmShopEditor.chkShow.Value = Shop(EditorIndex).ShowInfo
    frmShopEditor.chkSellsItems.Value = Shop(EditorIndex).BuysItems
    
    cItemMade = False
    
    frmShopEditor.cmbCurrency.Clear
    frmShopEditor.lstItems.Clear
    
    'Add all the currency items to cmbCurrency
    For i = 1 To MAX_ITEMS
        If item(i).Type = ITEM_TYPE_CURRENCY Then
            'It's a currency item, so add it to the list
            frmShopEditor.cmbCurrency.addItem (i & " - " & Trim(item(i).Name))
            'Add it to the item data so that we know the number
            frmShopEditor.cmbCurrency.ItemData(frmShopEditor.cmbCurrency.ListCount - 1) = i
            cItemMade = True 'we have at least 1 currency item
            If Shop(EditorIndex).currencyItem = i Then
                frmShopEditor.cmbCurrency.ListIndex = frmShopEditor.cmbCurrency.ListCount - 1
            End If
        End If
    Next i
    
    If Not cItemMade Then
        Call MsgBox("Please make at least one type of currency first!")
        Call ShopEditorCancel
        Exit Sub
    End If
    
    'Add all the items to the list
    For i = 1 To MAX_SHOP_ITEMS
        itemN = Shop(EditorIndex).ShopItem(i).ItemNum
    
        'If the item is not empty
        If itemN > 0 Then
            'Add the item to the shop list
            Call frmShopEditor.AddShopItem(itemN, Shop(EditorIndex).ShopItem(i).Price, Shop(EditorIndex).currencyItem, Shop(EditorIndex).ShopItem(i).Amount)
        End If
    Next i
    
    'Add all items to the 'add item' list
    For i = 1 To MAX_ITEMS
        frmShopEditor.cmbItemList.addItem (i & " - " & Trim(item(i).Name))
    Next i
    
    frmShopEditor.frmAddEditItem.Visible = False
    
    'Init shop editor temp array
    frmShopEditor.LoadShopItemData (EditorIndex)
    
    frmShopEditor.Show vbModal

   On Error GoTo 0
   Exit Sub

ShopEditorInit_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ShopEditorInit of Module modGameLogic"
    'Close the shop editor
    Unload frmShopEditor
    Call ShopEditorCancel
End Sub


Public Sub ShopEditorOk()
    Dim i As Integer
    Dim currencyItem As Integer
    
    If frmShopEditor.cmbCurrency.ListIndex < 0 Then
        MsgBox "Please pick a currency item!", vbExclamation
        Exit Sub
    End If
    
    currencyItem = frmShopEditor.cmbCurrency.ItemData(frmShopEditor.cmbCurrency.ListIndex)
    
    Shop(EditorIndex).Name = frmShopEditor.txtName.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.Value
    Shop(EditorIndex).BuysItems = frmShopEditor.chkSellsItems.Value
    Shop(EditorIndex).ShowInfo = frmShopEditor.chkShow.Value
    Shop(EditorIndex).currencyItem = currencyItem
    
    For i = 1 To MAX_SHOP_ITEMS
        Shop(EditorIndex).ShopItem(i).Amount = frmShopEditor.GetShopItemAmt(i)
        Shop(EditorIndex).ShopItem(i).ItemNum = frmShopEditor.GetShopItemNum(i)
        Shop(EditorIndex).ShopItem(i).Price = frmShopEditor.GetShopItemPrice(i)
    Next i
    
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

    frmSpellEditor.cmbClassReq.addItem "All Classes"
    For i = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.addItem Trim$(Class(i).Name)
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
        If item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
            frmPlayerTrade.PlayerInv1.addItem i & ": " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                frmPlayerTrade.PlayerInv1.addItem i & ": " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
            Else
                frmPlayerTrade.PlayerInv1.addItem i & ": " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).Name)
            End If
        End If
    Else
        frmPlayerTrade.PlayerInv1.addItem "<Nothing>"
    End If
Next i
    
    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub



Sub PlayerSearch(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Long, y1 As Long

    x1 = Int(X / PIC_X)
    y1 = Int(Y / PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        If Button = 1 Then
        frmMirage.picrclick.Visible = False
            Call MoveCharacter(Button, x1, y1)
            Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
        Else
        frmMirage.picrclick.Visible = False
        Dim i
            For i = 1 To MAX_PLAYERS
    If IsPlaying(i) = True Then
        If x1 = GetPlayerX(i) And y1 = GetPlayerY(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            RWINDEX = i
            frmMirage.lbrcname = GetPlayerName(i)
            frmMirage.lbrclvl = "lvl " & GetPlayerLevel(i) & " " & Trim$(Class(GetPlayerClass(i)).Name)
            frmMirage.picrclick.Top = Y + frmMirage.picUber.Top
            frmMirage.picrclick.Left = X + frmMirage.picUber.Left
            frmMirage.picrclick.Visible = True
            Exit Sub
        End If
    End If
Next i
        End If
    End If
    MouseDownX = x1
    MouseDownY = y1
End Sub

Sub BltTile2(ByVal X As Long, ByVal Y As Long, ByVal Tile As Long)
If TileFile(10) = 0 Then Exit Sub

    rec.Top = Int(Tile / TilesInSheets) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) + sx - NewXOffset, Y - (NewPlayerY * PIC_Y) + sx - NewYOffset, DD_TileSurf(10), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    'DisplayFx DD_TileSurf(10), (x - NewPlayerX * PIC_X) + sx - NewXOffset, y - (NewPlayerY * PIC_Y) + sx - NewYOffset, 32, 16, vbSrcAnd, DDBLT_ROP Or DDBLT_WAIT, Tile
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
    
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).xOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - Int(PIC_Y) + 75
    
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

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 4) - 4, 3)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 4) - 4, 4)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 15)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 4, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), QBColor(White))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub Bltscriptbubble(ByVal Index As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal Colour As Long)
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
    TextX = X * PIC_X - 22
    TextY = Y * PIC_Y - 22
    
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
        If strLine(intLoop) <> vbNullString Then
            Call DrawText(TexthDC, TextX + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * 8) \ 2) - 7, TextY, strLine(intLoop), QBColor(Colour))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub

Sub BltPlayerBars(ByVal Index As Long)
Dim X As Long, Y As Long

X = (GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
Y = (GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset


If Player(Index).HP = 0 Then Exit Sub
If Spritesize = 1 Then
'draws the back bars
   Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
   Call DD_BackBuffer.DrawBox(X, Y - 30, X + 32, Y - 34)
   
   'draws HP
   Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
   Call DD_BackBuffer.DrawBox(X, Y - 30, X + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), Y - 34)
Else
    If Spritesize = 2 Then
        'draws the back bars
           Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
           Call DD_BackBuffer.DrawBox(X, Y - 30 - PIC_Y, X + 32, Y - 34 - PIC_Y)
           
           'draws HP
           Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
           Call DD_BackBuffer.DrawBox(X, Y - 30 - PIC_Y, X + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), Y - 34 - PIC_Y)
    Else
        'draws the back bars
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 2, X + 32, Y - 2)
        
        'draws HP
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 2, X + ((Player(Index).HP / 100) / (Player(Index).MaxHp / 100) * 32), Y - 2)
    End If
End If
End Sub
Sub BltNpcBars(ByVal Index As Long)
Dim X As Long, Y As Long

   On Error GoTo BltNpcBars_Error

    If MapNpc(Index).HP = 0 Then Exit Sub
    If MapNpc(Index).num < 1 Then Exit Sub

    If Npc(MapNpc(Index).num).Big = 1 Then
        X = (MapNpc(Index).X * PIC_X + sx - 9 + MapNpc(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
        Y = (MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 32, X + 50, Y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        If MapNpc(Index).MaxHp < 1 Then
            Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHp + 1) / 100) * 50), Y + 36)
        Else
            Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 50), Y + 36)
        End If
    Else
        X = (MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset) - (NewPlayerX * PIC_X) - NewXOffset
        Y = (MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).yOffset) - (NewPlayerY * PIC_Y) - NewYOffset
        
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(X, Y + 32, X + 32, Y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
        
        If MapNpc(Index).MaxHp < 1 Then
            Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(Index).HP / 100) / ((MapNpc(Index).MaxHp + 1) / 100) * 32), Y + 36)
        Else
            Call DD_BackBuffer.DrawBox(X, Y + 32, X + ((MapNpc(Index).HP / 100) / (MapNpc(Index).MaxHp / 100) * 32), Y + 36)
        End If
        
    End If


   On Error GoTo 0
   Exit Sub

BltNpcBars_Error:

    If Err.Number = DDERR_CANTCREATEDC Then
        
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
        If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmMirage.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmMirage.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerLegsSlot(MyIndex) = Index Then Call BitBlt(frmMirage.LegsImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerRingSlot(MyIndex) = Index Then Call BitBlt(frmMirage.RingImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerNecklaceSlot(MyIndex) = Index Then Call BitBlt(frmMirage.NecklaceImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
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
            If Not item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
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

Public Sub UpdateotherVisInv()
Dim Index As Long
Dim d As Long

    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).num > 0 Then
            If Not item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
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

Sub BltSpriteChange(ByVal X As Long, ByVal Y As Long)
Dim x2 As Long, y2 As Long
    
    If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = Y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = X * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        If Spritesize = 0 Then
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data1 * PIC_Y + 16
        rec.Bottom = rec.Top + PIC_Y - 16
        Else
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data1 * 64 + 16
        rec.Bottom = rec.Top + 64 - 16
        End If
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = X * PIC_X + sx
        y2 = Y * PIC_Y + sx

                                       
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        If Spritesize = 1 Then
                Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y * 2) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    End If
End Sub

Sub BltSpriteChange2(ByVal X As Long, ByVal Y As Long)
Dim x2 As Long, y2 As Long
    
    If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_SPRITE_CHANGE Then

        With rec_pos
            .Top = Y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = X * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        If Spritesize = 0 Then
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data1 * PIC_Y + 16
        rec.Bottom = rec.Top + PIC_Y - 16
        Else
        rec.Top = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data1 * 64 + 16
        rec.Bottom = rec.Top + 64 - 16
        End If
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

Sub SendGameTime()
Dim packet As String

packet = "GmTime" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR
Call SendData(packet)
End Sub



Sub UpdateBank()
Dim i As Long

frmBank.lstInventory.Clear
frmBank.lstBank.Clear

For i = 1 To MAX_INV
If GetPlayerInvItemNum(MyIndex, i) > 0 Then
If item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
frmBank.lstInventory.addItem i & "> " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
Else
frmBank.lstInventory.addItem i & "> " & Trim$(item(GetPlayerInvItemNum(MyIndex, i)).Name)
End If
Else
frmBank.lstInventory.addItem i & "> Empty"
End If
Next i

For i = 1 To MAX_BANK
If GetPlayerBankItemNum(MyIndex, i) > 0 Then
If item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or item(GetPlayerBankItemNum(MyIndex, i)).Stackable = 1 Then
frmBank.lstBank.addItem i & "> " & Trim$(item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
Else
frmBank.lstBank.addItem i & "> " & Trim$(item(GetPlayerBankItemNum(MyIndex, i)).Name)
End If
Else
frmBank.lstBank.addItem i & "> Empty"
End If
Next i

frmBank.lstBank.ListIndex = 0
frmBank.lstInventory.ListIndex = 0
End Sub

Public Sub HouseEditorChooseTile(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        EditorTileX = Int(X / PIC_X)
        EditorTileY = Int(Y / PIC_Y)
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
    frmMapEditor.Visible = False
    frmMirage.Show
    frmHouseEditor.MousePointer = 1
    frmMirage.MousePointer = 1
    LoadMap (GetPlayerMap(MyIndex))
    'frmMirage.picMapEditor.Visible = False
End Sub

Public Sub CanonShoot(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim x1, y1 As Long

    x1 = Int(X / PIC_X)
    y1 = Int(Y / PIC_Y)
    Call SendData("canonshoot" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    
End Sub
        
Public Sub HouseEditorMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long, PicX As Long

    If InHouseEditor Then
        x1 = Int(X / PIC_X)
        y1 = Int(Y / PIC_Y)
        
        If frmHouseEditor.MousePointer = 2 Then
            If frmHouseEditor.mnuType(1).Checked = True Then
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
                   ElseIf frmHouseEditor.mnuType(2).Checked = True Then
                        With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                            If .Type = TILE_TYPE_WALKABLE Then
                                If frmMapEditor.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
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
                ElseIf frmHouseEditor.mnuType(2).Checked = True Then
                    With Map(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If .Type = TILE_TYPE_BLOCKED Then
                        .Type = 0
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .String1 = vbNullString
                        .String2 = vbNullString
                        .String3 = vbNullString
                    End If
                    End With
                End If
            End If
        End If
    End If
End Sub

'Sets the speed of a character based on speed
Sub SetSpeed(ByVal run As String, ByVal speed As Long)
    If LCase$(run) = "walk" Then
        SS_WALK_SPEED = speed
    ElseIf LCase$(run) = "run" Then
        SS_RUN_SPEED = speed
    End If
    'Ignore all other cases
End Sub

Sub MoveCharacter(ByVal Button As Long, ByVal MX As Integer, ByVal MY As Integer)
    If Player(MyIndex).input = 0 Then
        Exit Sub
    End If
            If GetPlayerY(MyIndex) = MAX_MAPY Then
                If MY = GetPlayerY(MyIndex) Then
                    Call warp
                End If
            Else
                If MY > GetPlayerY(MyIndex) And val(MY - GetPlayerY(MyIndex)) > val(MX - GetPlayerX(MyIndex)) Then
                        Call SetPlayerDir(MyIndex, 1)
                            If CanMove = True Then
                                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                                DirDown = True
                                Call SendPlayerMovemouse
                                Exit Sub
                            End If
                End If
            End If
            
            If GetPlayerY(MyIndex) = 0 Then
                If MY = GetPlayerY(MyIndex) Then
                    Call warp
                End If
            Else
                If MY < GetPlayerY(MyIndex) And val(MY - GetPlayerY(MyIndex)) < val(MX - GetPlayerX(MyIndex)) Then
                        Call SetPlayerDir(MyIndex, 0)
                            If CanMove = True Then
                                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                                DirUp = True
                                Call SendPlayerMovemouse
                                Exit Sub
                            End If
                End If
            End If
            
            If GetPlayerX(MyIndex) + 1 = MAX_MAPX Then
                If MX = GetPlayerX(MyIndex) Then
                    Call warp
                End If
            Else
                If MX > GetPlayerX(MyIndex) And val(MY - GetPlayerY(MyIndex)) < val(MX - GetPlayerX(MyIndex)) Then
                        Call SetPlayerDir(MyIndex, 3)
                            If CanMove = True Then
                                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                                DirRight = True
                                Call SendPlayerMovemouse
                                Exit Sub
                            End If
                End If

            End If
            
            If GetPlayerX(MyIndex) = 0 Then
                If MX = GetPlayerX(MyIndex) Then
                    Call warp
                End If
            Else
                If MX < GetPlayerX(MyIndex) And val(MY - GetPlayerY(MyIndex)) > val(MX - GetPlayerX(MyIndex)) Then
                        Call SetPlayerDir(MyIndex, 2)
                            If CanMove = True Then
                                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                                DirLeft = True
                                Call SendPlayerMovemouse
                                Exit Sub
                            End If
                End If
            End If
End Sub

Sub AlwaysOnTop(FormName As Form, bOnTop As Boolean)
    'Sets a form as always on top
    ' Bugfix - returns a long, not an integer. Pickle
    Dim Success As Long
    If bOnTop = False Then
        Success = SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    Else
        Success = SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, FLAGS)
    End If
End Sub

Sub GoShop(ByVal Shop As Integer)
    'Close any other shop windows
    frmNewShop.Hide
    
    'Initialize the shop
    Call frmNewShop.loadShop(Shop)
    snumber = Shop
    
    'Hide panel
    frmNewShop.picItemInfo.Visible = False
    
    'Show shop
    frmNewShop.Show vbModeless, frmMirage
    
    'Set focus
    frmNewShop.SetFocus
    
    'Show page 1 (it starts from 0)
    frmNewShop.showPage (0)
End Sub

Sub IncrementGameClock()
    Dim Time As String
    
    Seconds = Seconds + Gamespeed
    
    If Seconds > 59 Then
        Minutes = Minutes + 1
        Seconds = Seconds - 60
    End If
    If Minutes > 59 Then
        Hours = Hours + 1
        Minutes = 0
    End If
    If Hours > 24 Then
        Hours = 1
    End If
    
    If Hours > 12 Then
        Time = STR$(Hours - 12)
    Else
        Time = Hours
    End If
    
    If Minutes < 10 Then
        Time = Time & ":0" & Minutes
    Else
        Time = Time & ":" & Minutes
    End If
    If Seconds < 10 Then
        Time = Time & ":0" & Seconds
    Else
        Time = Time & ":" & Seconds
    End If
    If Hours > 12 Then
        Time = Time & " PM"
    Else
        Time = Time & " AM"
    End If
    
    frmMirage.GameClock.Caption = Time
        
End Sub
