Attribute VB_Name = "modGameLogic"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function TransparentBlt Lib "msimg32" ( _
    ByVal hdcDest As Long, ByVal iXStartDest As Long, ByVal iYStartDest As Long, _
    ByVal iWidthDest As Long, ByVal iHeightDest As Long, ByVal hdcSource As Long, _
    ByVal iXStartSrc As Long, ByVal iYStartSrc As Long, ByVal iWidthSrc As Long, _
    ByVal iHeightSrc As Long, ByVal iTransparentColour As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" ( _
    ByVal hdcDest As Long, ByVal iXStartDest As Long, ByVal iYStartDest As Long, _
    ByVal iWidthDest As Long, ByVal iHeightDest As Long, ByVal hdcSource As Long, _
    ByVal iXStartSrc As Long, ByVal iYStartSrc As Long, ByVal iWidthSrc As Long, _
    ByVal iHeightSrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal _
    hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal _
    lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'For Alpha Blending
Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, _
    ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
    ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)

Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

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

' encryption key
'secret key don't tell anyone.
Public Const EncryptionKey = "shers^&$ss esgegSHEHsh56 674&%$& wgsa"

' Speed moving vars
Public Const WALK_SPEED = 4
Public Const RUN_SPEED = 8

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

'Inventory/spell list
Public selectedInvItem As Long
Public selectedBankItem As Long
Public selectedBankInvLocalItem As Long

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
Public EditorMinLevel As Long
Public EditorMaxLevel As Long
Public EditorMsg As String
Public EditorDamage As Long
Public EditorNoDamage As Long
Public EditorSignNumber As Long
Public EditorTileSheet_Ground As Long
Public EditorTileSheet_Fringe As Long
Public EditorTileSheet_Anim As Long
Public EditorTileSheet_Mask As Long
Public EditorNPC_Num As Long
Public InBank As Boolean

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long

' Map for local use
Public SaveMap As MapRec
Public SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InSignEditor As Boolean
Public InPrayerEditor As Boolean
Public EditorIndex As Long
Public InquestEditor As Boolean
Public InWarp As Boolean

' Game fps
Public GameFPS As Long

' Used for atmosphere
Public GameWeather As Long
Public GameTime As Long

'sound
Public MP3 As New MP3Class
Public MP31 As New MP3Class
Public notInGame As Boolean


Sub Main()
Dim i As Long
Dim user As String
Dim pass As String
Dim dummy As String
Dim filenum As Long
    Call initColours
filenum = FreeFile

'    Open App.Path & "\setup.txt" For Input As #filenum
'        On Error Resume Next
'        Input #filenum, user, pass, dummy
'    Close #filenum
'    DoEvents
'    Open App.Path & "\setup.txt" For Output As #filenum
'        Print #filenum, user
'        Print #filenum, pass
'        Print #filenum, "version=" & App.Major & "." & App.Minor & "." & App.Revision
'    Close #filenum
    ' Check if the maps directory is there, if its not make it
    If LCase(Dir(App.Path & "\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\maps")
    End If
    
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InquestEditor = False
    InWarp = False
    InShopEditor = False
    MuteSound = False
    MuteMusic = False
    showLightning = False
    
    'setup the combo controlls
    
    frmMirage.cmbLayers.ListIndex = 0
    frmMirage.cmbAtributes.ListIndex = 0
    frmMirage.cmbTilePack.Enabled = True
    frmMirage.cmbLayers.Enabled = True
    frmMirage.cmbAtributes.Enabled = False

    
    ' Clear out players
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next i
    Call ClearTempTile
    
    frmSendGetData.Visible = True
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    Call SetStatus("Initializing FMod...")
    frmMainMenu.Visible = True
    frmSendGetData.Visible = False

    
End Sub

Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
End Sub

Sub MenuState(ByVal State As Long)
    frmSendGetData.Visible = True
    Call SetStatus("Connecting to server...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmNewAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmNewAccount.txtName.text, frmNewAccount.txtPassword.text)
            End If
            
        Case MENU_STATE_DELACCOUNT
            frmDeleteAccount.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending account deletion request ...")
                Call SendDelAccount(frmDeleteAccount.txtName.text, frmDeleteAccount.txtPassword.text)
            End If
        
        Case MENU_STATE_LOGIN
            frmLogin.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmLogin.txtName.text, frmLogin.txtPassword.text)
            End If
        
        Case MENU_STATE_NEWCHAR
            frmChars.Visible = False
            Call SetStatus("Connected, getting available classes...")
            Call SendGetClasses
            
        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character addition data...")
                If frmNewChar.optMale.value = True Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1, Val(frmNewChar.lblSpriteNo.Caption))
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1, Val(frmNewChar.lblSpriteNo.Caption))
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

    If Not IsConnected Then
        frmMainMenu.Visible = True
        frmSendGetData.Visible = False
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit http://afterdarkness.squiggleuk.com/", vbOKOnly, GAME_NAME)
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
Dim miniRec As RECT

    
    ' Set the focus
    'frmMirage.picScreen.SetFocus
    'frmMirage.txtSend.SetFocus
   ' MyText = MyText & frmMirage.txtSend
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
                
        ' Blit out tiles layers ground/anim1/anim2
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltTile(x, y)
            Next x
        Next y
                    
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call BltItem(i)
            End If
        Next i
        
        'blit out the players hp/mp bars
'        For i = 1 To MAX_PLAYERS
'            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
'                Call BltPlayerBars(i)
'            End If
'        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpc(i)
            
        Next i
        
        ' Blit out players
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayer(i)
                'Call BltPet(i)
            End If
        Next i
                
        ' Blit out tile layer fringe
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltFringeTile(x, y)
            Next x
        Next y
        Dim blackAmount As Long
        ' Blit out nightsky if its night time
        If ((blnNight = True) And map.Night = 0) Or map.Night = 1 Then
            'For y = 0 To MAX_MAPY
             '   For x = 0 To MAX_MAPX
            '        'Call BltTileNight(x, y)
            '    Next x
            'Next y
            If blackAmount <> 130 Then
                blackAmount = blackAmount + 5
                If blackAmount > 130 Then blackAmount = 130
            End If
        Else
            If blackAmount <> 0 Then
                blackAmount = blackAmount - 5
                If blackAmount < 0 Then blackAmount = 0
            End If
        End If
        
        ' Blit out nightsky for allways night places
'        If Map.Night = 1 Then
'            'For y = 0 To MAX_MAPY
'                'For x = 0 To MAX_MAPX
'                    'Call BltTileNight(x, y)
'                    'Dim test As Long
'                    'test = DD_BackBuffer.
'
'                'Next x
'            'Next y
'            blackAmount = 140
'        End If
                
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayerName(i)
            End If
        Next i
        
'        For i = 1 To MAX_PLAYERS
'            If Pets(i).map = GetPlayerMap(MyIndex) Then
'                Call BltPetName(i)
'            End If
'        Next i
        
        'draw night
            Call SquareAlphaBlend(frmMirage.picScreen.Width, frmMirage.picScreen.Height, frmMirage.picBlack.hdc, 0, 0, TexthDC, 0, 0, blackAmount)
        If showLightning And GameWeather = 1 And map.Night = 0 Then
        'draw lightning
            Call SquareAlphaBlend(frmMirage.picScreen.Width, frmMirage.picScreen.Height, frmMirage.picWhite.hdc, 0, 0, TexthDC, 0, 0, 130)
            showLightning = False
        End If
        ' Blit out attribs if in editor
        If InEditor Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    With map.Tile(x, y)
                        If .type = TILE_TYPE_BLOCKED Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "B", QBColor(BrightRed))
                        If .type = TILE_TYPE_WARP Or .type = TILE_TYPE_WARP_LEVEL Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "W", QBColor(BrightBlue))
                        If .type = TILE_TYPE_ITEM Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "I", QBColor(15))
                        If .type = TILE_TYPE_NPCAVOID Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "N", QBColor(15))
                        If .type = TILE_TYPE_KEY Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "K", QBColor(15))
                        If .type = TILE_TYPE_KEYOPEN Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "O", QBColor(15))
                        If .type = TILE_TYPE_DAMAGE Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "D", QBColor(BrightRed))
                        If .type = TILE_TYPE_HEAL Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "H", QBColor(BrightGreen))
                        If .type = TILE_TYPE_SIGN Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "S", QBColor(Blue))
                        If .type = TILE_TYPE_LEVEL Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "L", QBColor(Yellow))
                        If .type = TILE_TYPE_NPC_SPAWN Then Call DrawText(TexthDC, x * PIC_X + 7, y * PIC_Y + 8, "NS", QBColor(BrightGreen))
                    End With
                Next x
            Next y
            
        End If
        
        ' Blit the text they are putting in
        'eDiT1281
        'frmMirage.txtSend.Text = frmMirage.txtSend.Text & MyText
        'Call DrawText(TexthDC, 0, (MAX_MAPY + 1) * PIC_Y - 20, MyText, RGB(255, 255, 255))
        ' Draw map name
        If map.Moral = MAP_MORAL_NONE Then
            frmMirage.lblMapName.Caption = Trim(map.Name)
            frmMirage.lblstreetname.Caption = Trim(map.street)
            'Call DrawText(TexthDC, Int((MAX_MAPX + 1) * PIC_X / 2) - (Int(Len(Trim(Map.name)) / 2) * 8), 1, Trim(Map.name), QBColor(15))
        ElseIf map.Moral = MAP_MORAL_SAFE Then
            frmMirage.lblMapName.Caption = Trim(map.Name) & "(S)"
            frmMirage.lblstreetname.Caption = Trim(map.street)
            'frmMirage.lblMapName.Caption = Trim(Replace(Map.name, "/n", vbCrLf)) & " (S)"
            'Call DrawText(TexthDC, Int((MAX_MAPX + 1) * PIC_X / 2) - (Int(Len(Trim(Map.name)) / 2) * 8), 1, Trim(Map.name), QBColor(15))
        ElseIf map.Moral = MAP_MORAL_ARENA Then
            frmMirage.lblMapName.Caption = Trim(map.Name) & "(A)"
            frmMirage.lblstreetname.Caption = Trim(map.street)
            'frmMirage.lblMapName.Caption = Trim(Replace(Map.name, "/n", vbCrLf)) & " (A)"
            'Call DrawText(TexthDC, Int((MAX_MAPX + 1) * PIC_X / 2) - (Int(Len(Trim(Map.name)) / 2) * 8), 1, Trim(Map.name), QBColor(7))
        ElseIf map.Moral = MAP_MORAL_SAVAGE Then
            frmMirage.lblMapName.Caption = Trim(map.Name) & "(SG)"
            frmMirage.lblstreetname.Caption = Trim(map.street)
            'frmMirage.lblMapName.Caption = Trim(Replace(Map.name, "/n", vbCrLf)) & " (SG)"
            'Call DrawText(TexthDC, Int((MAX_MAPX + 1) * PIC_X / 2) - (Int(Len(Trim(Map.name)) / 2) * 8), 1, Trim(Map.name), RGB(229, 78, 67))
        End If
        
        ' Check if we are getting a map, and if we are tell them so
        'If GettingMap = True Then
        '    Call DrawText(TexthDC, 50, 50, "Receiving Map...", QBColor(BrightCyan))
        'End If
                        
        ' Release DC
        Call DD_BackBuffer.ReleaseDC(TexthDC)
        
        ' Get the rect for the back buffer to blit from
        rec.top = 0
        rec.Bottom = (MAX_MAPY + 1) * PIC_Y
        rec.Left = 0
        rec.Right = (MAX_MAPX + 1) * PIC_X
        
        ' Get the rect to blit to
        Call dx.GetWindowRect(frmMirage.picScreen.hwnd, rec_pos)
        rec_pos.Bottom = rec_pos.top + ((MAX_MAPY + 1) * PIC_Y)
        rec_pos.Right = rec_pos.Left + ((MAX_MAPX + 1) * PIC_X)
        
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
                'Call ProcessPetMovement(i)
            End If
        Next i
        
        ' Process npc movements (actually move them)
        For i = 1 To MAX_MAP_NPCS
            If map.Npc(i) > 0 Then
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
        Do While GetTickCount < Tick + 50
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
    MP3.MP3Stop
    End
End Sub

Sub BltTileNight(ByVal x As Long, ByVal y As Long, Optional ByVal colour As Long = 0)
    'Night code
'    Dim i As Long
'    Dim j As Long
'    Dim myRect As RECT
'    For i = 0 To PIC_Y Step 2
'        For j = 0 To PIC_X Step 2
'            myRect.top = (y * PIC_Y) + j
'            myRect.Bottom = myRect.top + 1
'            myRect.Left = (x * PIC_X) + i
'            myRect.Right = myRect.Left + 1
'            Call DD_BackBuffer.BltColorFill(myRect, colour)
'        Next j
'    Next i
    
    
    'NEW NEW
    'Call SquareAlphaBlend(frmMirage.picScreen.Width, frmMirage.picScreen.Height, DD_BackBuffer.GetDC, 0, 0, frmMirage.picBlack, 0, 0, 100)
    
    
    'Call TransparentBlt(DD_BackBuffer.BltColorFill(, X * PIC_X, Y * PIC_Y, PIC_X, PIC_Y, frmMirage.picNight.hdc, 32, 0, 32, 32, RGB(0, 0, 0))
    'Call BitBlt(frmItemEditor.picPic.hdc, 0, 0, PIC_X, PIC_Y, frmItemEditor.picItems.hdc, 0, frmItemEditor.scrlPic.Value * PIC_Y, SRCCOPY)
    'rec.top = 0
    'rec.Bottom = 32
    'rec.Left = 0
    'rec.Right = 32
    
    'rec.top = 0
    'rec.Bottom = 32
    'rec.Left = 32
    'rec.Right = 64
    'Call DD_BackBuffer.BltFast(X * PIC_X, Y * PIC_Y, DD_NIGHTSURF, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltTile(ByVal x As Long, ByVal y As Long)
Dim Ground As Long
Dim Anim1 As Long
Dim Anim2 As Long
Dim TileSheet_Ground As Long
Dim TileSheet_Anim As Long
Dim TileSheet_Mask As Long

    Ground = map.Tile(x, y).Ground
    Anim1 = map.Tile(x, y).mask
    Anim2 = map.Tile(x, y).Anim
    TileSheet_Ground = map.Tile(x, y).TileSheet_Ground
    TileSheet_Mask = map.Tile(x, y).TileSheet_Mask
    TileSheet_Anim = map.Tile(x, y).TileSheet_Anim
    If TileSheet_Ground >= MAX_TILE_SHEETS Then TileSheet_Ground = 0
    If TileSheet_Mask >= MAX_TILE_SHEETS Then TileSheet_Mask = 0
    If TileSheet_Anim >= MAX_TILE_SHEETS Then TileSheet_Anim = 0
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    rec.top = Int(Ground / 50) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (Ground - Int(Ground / 50) * 50) * PIC_X
    rec.Right = rec.Left + PIC_X
    'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT)
    Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf(TileSheet_Ground), rec, DDBLTFAST_WAIT)
    
    If (MapAnim = 0) Or (Anim2 <= 0) Then
        ' Is there an animation tile to plot?
        If Anim1 > 0 And TempTile(x, y).DoorOpen = NO Then
            rec.top = Int(Anim1 / 50) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Anim1 - Int(Anim1 / 50) * 50) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf(TileSheet), rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf(TileSheet_Mask), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim2 > 0 Then
            rec.top = Int(Anim2 / 50) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Anim2 - Int(Anim2 / 50) * 50) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf(TileSheet_Anim), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = MapItem(ItemNum).y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = MapItem(ItemNum).x * PIC_X
        .Right = .Left + PIC_X
    End With

    rec.top = Item(MapItem(ItemNum).num).Pic * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = 0
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(MapItem(ItemNum).x * PIC_X, MapItem(ItemNum).y * PIC_Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub blitEquip()
Dim helm As Long
Dim armor As Long
Dim shield As Long
Dim weapon As Long
helm = Player(MyIndex).HelmetSlot
weapon = Player(MyIndex).WeaponSlot
armor = Player(MyIndex).ArmorSlot
shield = Player(MyIndex).ShieldSlot

If weapon = 0 Then
    weapon = 0
Else
    weapon = Item(GetPlayerInvItemNum(MyIndex, weapon)).Pic
End If
If armor = 0 Then
    armor = 0
Else
    armor = Item(GetPlayerInvItemNum(MyIndex, armor)).Pic
End If
If helm = 0 Then
    helm = 0
Else
    helm = Item(GetPlayerInvItemNum(MyIndex, helm)).Pic
End If
If shield = 0 Then
    shield = 0
Else
    shield = Item(GetPlayerInvItemNum(MyIndex, shield)).Pic
End If

'weapon = Item(GetPlayerInvItemNum(MyIndex, weapon)).Pic
'armor = Item(GetPlayerInvItemNum(MyIndex, armor)).Pic
'helm = Item(GetPlayerInvItemNum(MyIndex, helm)).Pic
'shield = Item(GetPlayerInvItemNum(MyIndex, shield)).Pic
'Debug.Print weapon
'Debug.Print armor
'Debug.Print helm
'Debug.Print shield
'Debug.Print "OK"
DoEvents
Call StretchBlt(frmMirage.picEquip(1).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, weapon * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
'Call BitBlt(frmMirage.picEquip(1).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, weapon)).Pic * PIC_Y, SRCCOPY)
DoEvents
'Call BitBlt(frmMirage.picEquip(3).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, armor)).Pic * PIC_Y, SRCCOPY)
Call StretchBlt(frmMirage.picEquip(3).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, armor * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
DoEvents
'Call BitBlt(frmMirage.picEquip(2).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, shield)).Pic * PIC_Y, SRCCOPY)
Call StretchBlt(frmMirage.picEquip(2).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, shield * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
DoEvents
'Call BitBlt(frmMirage.picEquip(0).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, helm)).Pic * PIC_Y, SRCCOPY)
Call StretchBlt(frmMirage.picEquip(0).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, helm * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
DoEvents
End Sub

'Public Sub blitMiniMap(ByVal x As Long, ByVal y As Long)
'Dim Ground As Long
'Dim Fringe As Long
'Dim mask As Long
'Dim TileSheet_Ground As Long
'Dim TileSheet_Fringe As Long
'Dim TileSheet_Mask As Long
'Dim rec As RECT
'
'Dim groundHDC As Long
'Dim maskHDC As Long
'Dim fringeHDC As Long
'
'
'
'    Ground = Map.Tile(x, y).Ground
'    mask = Map.Tile(x, y).mask
'    Fringe = Map.Tile(x, y).Fringe
'    TileSheet_Ground = Map.Tile(x, y).TileSheet_Ground
'    TileSheet_Mask = Map.Tile(x, y).TileSheet_Mask
'    TileSheet_Fringe = Map.Tile(x, y).TileSheet_Fringe
'    If TileSheet_Ground >= MAX_TILE_SHEETS Then TileSheet_Ground = 0
'    If TileSheet_Mask >= MAX_TILE_SHEETS Then TileSheet_Mask = 0
'    If TileSheet_Fringe >= MAX_TILE_SHEETS Then TileSheet_Fringe = 0
'
'    rec.top = Int(Ground / 50) * PIC_Y
'    rec.Bottom = rec.top + PIC_Y
'    rec.Left = (Ground - Int(Ground / 50) * 50) * PIC_X
'    rec.Right = rec.Left + PIC_X
'
'
'    groundHDC = DD_TileSurf_MINI(TileSheet_Ground).GetDC
'    Call StretchBlt(frmMirage.picMiniMap.hdc, (x * 11), (y * 11), 11, 11, groundHDC, rec.Left, rec.top, PIC_X, PIC_Y, SRCCOPY)
'    DD_TileSurf_MINI(TileSheet_Ground).ReleaseDC (groundHDC)
'
'
'    If mask > 0 Then
'        rec.top = Int(mask / 50) * PIC_Y
'        rec.Bottom = rec.top + PIC_Y
'        rec.Left = (mask - Int(mask / 50) * 50) * PIC_X
'        rec.Right = rec.Left + PIC_X
'
'        maskHDC = DD_TileSurf_MINI(TileSheet_Mask).GetDC
'        Call TransparentBlt(frmMirage.picMiniMap.hdc, (x * 11), (y * 11), 11, 11, maskHDC, rec.Left, rec.top, PIC_X, PIC_Y, RGB(0, 0, 0))
'        DD_TileSurf_MINI(TileSheet_Mask).ReleaseDC (maskHDC)
'    End If
'
'    If Fringe > 0 Then
'        rec.top = Int(Fringe / 50) * PIC_Y
'        rec.Bottom = rec.top + PIC_Y
'        rec.Left = (Fringe - Int(Fringe / 50) * 50) * PIC_X
'        rec.Right = rec.Left + PIC_X
'
'        fringeHDC = DD_TileSurf_MINI(TileSheet_Fringe).GetDC
'        Call TransparentBlt(frmMirage.picMiniMap.hdc, (x * 11), (y * 11), 11, 11, fringeHDC, rec.Left, rec.top, PIC_X, PIC_Y, RGB(0, 0, 0))
'        DD_TileSurf_MINI(TileSheet_Fringe).ReleaseDC (fringeHDC)
'    End If
    
'    Call StretchBlt(frmMirage.picEquip(1).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, weapon * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
'    'Call BitBlt(frmMirage.picEquip(1).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, weapon)).Pic * PIC_Y, SRCCOPY)
'    DoEvents
'    'Call BitBlt(frmMirage.picEquip(3).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, armor)).Pic * PIC_Y, SRCCOPY)
'    Call StretchBlt(frmMirage.picEquip(3).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, armor * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
'    DoEvents
'    'Call BitBlt(frmMirage.picEquip(2).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, shield)).Pic * PIC_Y, SRCCOPY)
'    Call StretchBlt(frmMirage.picEquip(2).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, shield * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
'    DoEvents
'    'Call BitBlt(frmMirage.picEquip(0).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, helm)).Pic * PIC_Y, SRCCOPY)
'    Call StretchBlt(frmMirage.picEquip(0).hdc, 0, 0, 42, 42, frmMirage.picItems.hdc, 0, helm * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
'    DoEvents
'End Sub

Sub BltFringeTile(ByVal x As Long, ByVal y As Long)
Dim Fringe As Long
Dim TileSheet As Long
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Fringe = map.Tile(x, y).Fringe
    TileSheet = map.Tile(x, y).TileSheet_Fringe
    If TileSheet > MAX_TILE_SHEETS - 1 Then TileSheet = 0
    If Fringe > 0 Then
        rec.top = Int(Fringe / 50) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (Fringe - Int(Fringe / 50) * 50) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf(TileSheet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltPlayerBars(ByVal Index As Long)
    Dim x As Long, y As Long, lngHP As Long, lngMaxHP As Long
    lngHP = GetPlayerHP(Index)
    lngMaxHP = GetPlayerMaxHP(Index)
    If lngMaxHP = 0 Then lngMaxHP = 1
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
    'Debug.Print "X" & x
    'Debug.Print "Y" & y
    
    'draw box for bars
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
    Call DD_BackBuffer.DrawBox(x, y + 35, x + 32, y + 39)
    
    'draw HP
    'Debug.Print "HP: " & GetPlayerHP(index)
    'Debug.Print "HPM: " & GetPlayerMaxHP(index)
    'Debug.Print "PER" & ((GetPlayerHP(index) / GetPlayerMaxHP(index)) * 32)
    Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
    Call DD_BackBuffer.DrawBox(x, y + 32, x + ((lngHP / lngMaxHP) * 32), y + 36)
    'draw MP
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    'Call DD_BackBuffer.DrawBox(x, y + 35, x + (GetPlayerMP(index) / GetPlayerMaxMP(index) * 32), y + 39)
  
    
End Sub


Sub BltNpcBar(ByVal Index As Long)
    Dim x As Long, y As Long
    'Debug.Print "HP: " & index & " " & MapNpc(index).HP
    'Debug.Print "HPM: " & MapNpc(index).MaxHP
    x = MapNpc(Index).x * PIC_X + MapNpc(Index).XOffset
    y = MapNpc(Index).y * PIC_Y + MapNpc(Index).YOffset - 4
    If MapNpc(Index).maxHP > 0 Then
        Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
        Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
        Call DD_BackBuffer.SetFillColor(RGB(0, 0, 255))
        If MapNpc(Index).HP <> 0 Then
            Call DD_BackBuffer.DrawBox(x, y + 32, x + ((MapNpc(Index).HP / MapNpc(Index).maxHP) * 32), y + 36)
        Else
            Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
        End If
    End If
    

End Sub



Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
        .Bottom = .top + PIC_Y
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
    
    rec.top = GetPlayerSprite(Index) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
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
    Color = GetPlayerColour(Index)
    If Color = 0 Then Color = 15
    If Color <= 15 And Color >= 0 Then Color = QBColor(Color)
    ' Draw name
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - Int(PIC_Y / 2) - 4
    'Debug.Print GetPlayerColour(index)
    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(Index), Color)
    
End Sub

'Sub BltPetName(ByVal petID As Long)
'Dim TextX As Long
'Dim TextY As Long
'Dim Color As Long
'    Color = QBColor(11)
'    ' Draw name
'    TextX = Pets(petID).x * PIC_X + Pets(petID).XOffset + Int(PIC_X / 2) - ((Len(Pets(petID).Name) / 2) * 8)
'    TextY = Pets(petID).y * PIC_Y + Pets(petID).YOffset - Int(PIC_Y / 2) - 4
'    'Debug.Print GetPlayerColour(index)
'    Call DrawText(TexthDC, TextX, TextY, Pets(petID).Name, Color)
'End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then
        Exit Sub
    End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .top + PIC_Y
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
    
    rec.top = Npc(MapNpc(MapNpcNum).num).sprite * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset
    y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.top = rec.top + (y * -1)
    End If
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    
    'blit out NPC bars
    Call BltNpcBar(MapNpcNum)
End Sub

'Sub BltPet(ByVal petID As Long)
'Dim Anim As Byte
'Dim x As Long, y As Long
'
'    If Pets(petID).map <> GetPlayerMap(MyIndex) Then
'        Exit Sub
'    End If
'
'    ' Check for animation
'    Anim = 0
'        Select Case Pets(petID).Dir
'            Case DIR_UP
'                If (Pets(petID).YOffset < PIC_Y / 2) Then Anim = 1
'            Case DIR_DOWN
'                If (Pets(petID).YOffset < PIC_Y / 2 * -1) Then Anim = 1
'            Case DIR_LEFT
'                If (Pets(petID).XOffset < PIC_Y / 2) Then Anim = 1
'            Case DIR_RIGHT
'                If (Pets(petID).XOffset < PIC_Y / 2 * -1) Then Anim = 1
'        End Select
'
'    rec.top = Pets(petID).sprite * PIC_Y
'    rec.Bottom = rec.top + PIC_Y
'    rec.Left = (Pets(petID).Dir * 3 + Anim) * PIC_X
'    rec.Right = rec.Left + PIC_X
'
'    x = Pets(petID).x * PIC_X + Pets(petID).XOffset
'    y = Pets(petID).y * PIC_Y + Pets(petID).YOffset - 4
'
'    ' Check if its out of bounds because of the offset
'    If y < 0 Then
'        y = 0
'        rec.top = rec.top + (y * -1)
'    End If
'
'    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
'End Sub

Sub ProcessMovement(ByVal Index As Long)
'STAMINA
'If Index = MyIndex Then
'    If GetPlayerSP(Index) < 1 Then
'        If Player(Index).Moving = MOVING_RUNNING Then
'            Player(Index).Moving = MOVING_WALKING
'        End If
'    End If
'End If
    ' Check if player is walking, and if so process moving them over
    If Player(Index).moving = MOVING_WALKING Then
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
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then
            Player(Index).moving = 0
        End If
    End If

    ' Check if player is running, and if so process moving them over
    If Player(Index).moving = MOVING_RUNNING Then
'        Call SetPlayerSP(Index, GetPlayerSP(Index) - 1)
'        Call SendStamina
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
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then
            Player(Index).moving = 0
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if player is walking, and if so process moving them over
    If MapNpc(MapNpcNum).moving = MOVING_WALKING Then
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
            MapNpc(MapNpcNum).moving = 0
        End If
    End If
End Sub

'Sub ProcessPetMovement(ByVal petID As Long)
'    ' Check if player is walking, and if so process moving them over
'    If Pets(petID).moving = MOVING_WALKING Then
'        Select Case Pets(petID).Dir
'            Case DIR_UP
'                Pets(petID).YOffset = Pets(petID).YOffset - WALK_SPEED
'            Case DIR_DOWN
'                Pets(petID).YOffset = Pets(petID).YOffset + WALK_SPEED
'            Case DIR_LEFT
'                Pets(petID).XOffset = Pets(petID).XOffset - WALK_SPEED
'            Case DIR_RIGHT
'                Pets(petID).XOffset = Pets(petID).XOffset + WALK_SPEED
'        End Select
'
'        ' Check if completed walking over to the next tile
'        If (Pets(petID).XOffset = 0) And (Pets(petID).YOffset = 0) Then
'            Pets(petID).moving = 0
'        End If
'    End If
'End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim JailNo() As String
Dim i As Long
Dim n As Long

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
        ' Broadcast message
        If Mid(MyText, 1, 1) = "'" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        ' Emote message
        If Mid(MyText, 1, 1) = "-" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = ""
            frmMirage.txtSend.text = MyText
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
                'Call AddText("Usage: !playername msghere", AlertColor)
                Call AddTextNew(frmMirage.txtChannelAll, "Usage: !playername msghere", RGB_AlertColor)
            End If
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
            
        ' // Commands //
        ' Help
        If LCase(Mid(MyText, 1, 5)) = "/help" Then
            Call AddTextNew(frmMirage.txtChannelAll, "Social Commands:", RGB_HelpColor)
            Call AddTextNew(frmMirage.txtChannelAll, "'msghere = Broadcast Message", RGB_HelpColor)
            Call AddTextNew(frmMirage.txtChannelAll, "-msghere = Emote Message", RGB_HelpColor)
            Call AddTextNew(frmMirage.txtChannelAll, "!namehere msghere = Player Message", RGB_HelpColor)
            Call AddTextNew(frmMirage.txtChannelAll, "Available Commands: /help, /who, /fps, /inv, /stats, /train, /trade, /party, /join, /leave", RGB_HelpColor)
            
'            Call AddText("Social Commands:", HelpColor)
'            Call AddText("'msghere = Broadcast Message", HelpColor)
'            Call AddText("-msghere = Emote Message", HelpColor)
'            Call AddText("!namehere msghere = Player Message", HelpColor)
'            Call AddText("Available Commands: /help, /info, /who, /fps, /inv, /stats, /train, /trade, /party, /join, /leave", HelpColor)
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        
        
        ' Whos Online
        If LCase(Mid(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
                        
        ' Checking fps
        If LCase(Mid(MyText, 1, 4)) = "/fps" Then
            Call AddTextNew(frmMirage.txtChannelAll, "FPS: " & GameFPS, RGB_HelpColor)
            'Call AddText("FPS: " & GameFPS, Pink)
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
                
        ' Show inventory
        If LCase(Mid(MyText, 1, 4)) = "/inv" Then
            Call UpdateInventory
            frmMirage.picInv.Visible = True
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        ' Request stats
        If LCase(Mid(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
    
        ' Show training
        If LCase(Mid(MyText, 1, 6)) = "/train" Then
            frmTraining.Show vbModal
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If

        ' Request stats
        If LCase(Mid(MyText, 1, 6)) = "/trade" Then
            Call SendData("trade" & SEP_CHAR & END_CHAR)
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        ' Request stats
        If LCase(Mid(MyText, 1, 4)) = "/bio" Then
            Call SendData("bio" & SEP_CHAR & Mid(MyText, 5) & SEP_CHAR & END_CHAR)
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        ' Party request
        If LCase(Mid(MyText, 1, 6)) = "/party" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                'Call AddText("Usage: /party playernamehere", AlertColor)
                Call AddTextNew(frmMirage.txtChannelAll, "Usage: /party playernamehere", RGB_AlertColor)
            End If
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        ' Join party
        If LCase(Mid(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        ' Leave party
        If LCase(Mid(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        ' item lib
        If LCase(Mid(MyText, 1, 8)) = "/itemlib" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 9 Then
                ChatText = Mid(MyText, 10, Len(MyText) - 9)
                Call SendItemLibRequest(ChatText)
            Else
                'Call AddText("Usage: /itemlib itemnamehere", AlertColor)
                Call AddTextNew(frmMirage.txtChannelAll, "Usage: /itemlib itemnamehere", RGB_AlertColor)
            End If
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
         ' exit game
        If LCase(Mid(MyText, 1, 6)) = "/exit" Then
                Call GameDestroy
            Exit Sub
        End If
         ' Quest
        If Mid(MyText, 1, 9) = "/quest" Then
            Call SendRequestQuest
            MyText = ""
            frmMirage.txtSend.text = MyText
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' Verification User
            If LCase(Mid(MyText, 1, 5)) = "/info" Then
                ChatText = Mid(MyText, 6, Len(MyText) - 5)
                Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
        
            ' Admin Help
            If LCase(Mid(MyText, 1, 6)) = "/admin" Then
                Call SendRequestAdminHelp
                'Call AddText("Social Commands:", HelpColor)
                'Call AddText("""msghere = Global Admin Message", HelpColor)
                'Call AddText("=msghere = Private Admin Message", HelpColor)
                'Call AddText("Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /jail, /edititem, /respawn, /editnpc, /motd, /editshop, /ban, /editspell", HelpColor)
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Kicking a player
            If LCase(Mid(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Jailing a player
            If LCase(Mid(MyText, 1, 5)) = "/jail" Then
            
                If Len(MyText) > 6 Then
                    JailNo = Split(Mid(MyText, 7, Len(MyText) - 6), " ")
                    'MyText = Mid(MyText, 7, Len(MyText) - 6)
                    If UBound(JailNo) >= 1 Then
                        Call SendJail(JailNo(0), JailNo(1))
                    Else
                        Call SendJail(JailNo(0))
                    End If
                End If
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
        
            ' Global Message
            If Mid(MyText, 1, 1) = """" Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Global Message
            If Mid(MyText, 1, 3) = "/m " Then
                ChatText = Mid(MyText, 4, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call MassMsg(ChatText)
                End If
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
        
            ' Admin Message
            If Mid(MyText, 1, 1) = "=" Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
        End If
        
        ' Warping to a player
            If LCase(Mid(MyText, 1, 9)) = "/warpmeto" Then
                If Len(MyText) > 10 Then
                    MyText = Mid(MyText, 10, Len(MyText) - 9)
                    Call WarpMeTo(MyText)
                End If
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
                        
            ' Warping a player to you
            If LCase(Mid(MyText, 1, 9)) = "/warptome" Then
                If Len(MyText) > 10 Then
                    MyText = Mid(MyText, 10, Len(MyText) - 9)
                    Call WarpToMe(MyText)
                End If
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
                        
            ' Warping to a map
            If LCase(Mid(MyText, 1, 7)) = "/warpto" Then
                If Len(MyText) > 8 Then
                    MyText = Mid(MyText, 8, Len(MyText) - 7)
                    Dim warparr() As String
                    warparr = Split(MyText, " ")
                    If UBound(warparr) = 3 Then
                        If Val(warparr(2)) > 15 Then warparr(2) = 15
                        If Val(warparr(3)) > 14 Then warparr(3) = 11
                        If Val(warparr(2)) < 0 Then warparr(2) = 0
                        If Val(warparr(3)) < 0 Then warparr(3) = 0
                        Call WarpTo(n, warparr(1), warparr(2), warparr(3))
                    Else
                        n = Val(MyText)

                        ' Check to make sure its a valid map #
                        If n > 0 And n <= MAX_MAPS Then
                            Call WarpTo(n)
                        Else
                            'Call AddText("Invalid map number.", Red)
                            Call AddTextNew(frmMirage.txtChannelAll, "Invalid map number.", RGB_AlertColor)
                        End If
                    End If
                End If
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If LCase(Mid(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                MyText = ""
                frmMirage.txtSend.text = MyText

                Exit Sub
            End If
            
            ' Map Editor
            If LCase(Mid(MyText, 1, 10)) = "/mapeditor" Then
            Dim t As Long
                Call SendRequestEditMap
                MyText = ""
                frmMirage.txtSend.text = MyText
                                
                If frmMirage.Width = FORM_X Then
                    For t = FORM_X To FORM_EDITOR_X Step 100
                        frmMirage.Width = t
                        frmMirage.Refresh
                        DoEvents
                    Next t
                Else
                    frmMirage.Width = FORM_EDITOR_X
                End If
                Exit Sub
            End If
            
            
            
            ' Setting sprite
            If LCase(Mid(MyText, 1, 10)) = "/setsprite" Then
                Load frmSprite
                frmSprite.Show
                frmSprite.SetFocus
                DoEvents
                MyText = ""
                frmMirage.txtSend.text = MyText
                Load frmSprite
                frmSprite.Show
                frmSprite.SetFocus
                DoEvents
                Exit Sub
            End If
            
            ' Map report
            If LCase(Mid(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
        
            ' Respawn request
            If Mid(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                MyText = ""
                frmMirage.txtSend.text = MyText
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
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Check the ban list
            If Mid(MyText, 1, 3) = "/banlist" Then
                Call SendBanList
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Banning a player
            If LCase(Mid(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    MyText = ""
                    frmMirage.txtSend.text = MyText
                End If
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Editing item request
            If Mid(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            'edit quest request
            If Mid(MyText, 1, 10) = "/editquest" Then
                Call SendRequestEditQuest
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
           
            
            ' Editing npc request
            If Mid(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNpc
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Editing shop request
            If Mid(MyText, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
        
            ' Editing spell request
            If Mid(MyText, 1, 10) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Editing prayer request
            If Mid(MyText, 1, 11) = "/editprayer" Then
                Call SendRequestEditPrayer
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Editing sign request
            If Mid(MyText, 1, 9) = "/editsign" Then
                Call SendRequestEditSign
                MyText = ""
                frmMirage.txtSend.text = MyText
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
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase(Mid(MyText, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                MyText = ""
                frmMirage.txtSend.text = MyText
                Exit Sub
            End If
        End If
        
        ' Say message
        If Len(Trim(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        MyText = ""
        frmMirage.txtSend.text = MyText
        'Exit Sub
    End If
    
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
            'frmMirage.txtSend.Text = frmMirage.txtSend.Text & Chr(KeyAscii)
            'MyText = MyText & Chr(KeyAscii)
        End If
    End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim(MyText) = "" Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & SEP_CHAR & END_CHAR & "mapgetsign" & SEP_CHAR & END_CHAR & "mapgetlevel" & SEP_CHAR & END_CHAR)
        'map sign on enter
        'Call SendData("mapgetsign" & SEP_CHAR & END_CHAR)
        'Call SendData("mapgetlevel" & SEP_CHAR & END_CHAR)
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
            If KeyCode = vbKeyF1 Then
                ' Map editor
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                    Dim t As Long
                    Call frmMirage.clearPanes
                    Call SendRequestEditMap
                    MyText = ""
                    frmMirage.txtSend.text = MyText
                                    
                    If frmMirage.Width = FORM_X Then
                        For t = FORM_X To FORM_EDITOR_X Step 100
                            frmMirage.Width = t
                            frmMirage.Refresh
                            DoEvents
                        Next t
                    Else
                        frmMirage.Width = FORM_EDITOR_X
                    End If
                End If
            End If
            If KeyCode = vbKeyF2 Then
                If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                    ' Editing item request
                        Call SendRequestEditItem
                        MyText = ""
                        frmMirage.txtSend.text = MyText
                        Exit Sub
                End If
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

Function canMove() As Boolean
Dim i As Long, d As Long

    canMove = True
    
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).moving <> 0 Then
        canMove = False
        Exit Function
    End If
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            canMove = False
            Exit Function
        End If
    End If
    
    d = GetPlayerDir(MyIndex)
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).type = TILE_TYPE_BLOCKED Then
                canMove = False
                    If GetPlayerY(MyIndex) - 1 >= 0 Then
                        Call WarpTo_U(GetPlayerMap(MyIndex), GetPlayerMap(MyIndex), GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1)
                    End If
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).type = TILE_TYPE_KEY Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).DoorOpen = NO Then
                    canMove = False
                    
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
                            canMove = False
                        
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
                        canMove = False
                        
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
            If map.Up > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            canMove = False
            Exit Function
        End If
    End If
            
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            ' Check to see if the map tile is blocked or not
            If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).type = TILE_TYPE_BLOCKED Then
                canMove = False
                    If GetPlayerY(MyIndex) + 1 <= 11 Then
                        Call WarpTo_U(GetPlayerMap(MyIndex), GetPlayerMap(MyIndex), GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1)
                    End If
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).type = TILE_TYPE_KEY Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).DoorOpen = NO Then
                    canMove = False
                
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
                        canMove = False
                        
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
                        canMove = False
                        
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
            If map.Down > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            canMove = False
            Exit Function
        End If
    End If
                
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            ' Check to see if the map tile is blocked or not
            If map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).type = TILE_TYPE_BLOCKED Then
                canMove = False
                    If GetPlayerX(MyIndex) - 1 >= 0 Then
                        Call WarpTo_U(GetPlayerMap(MyIndex), GetPlayerMap(MyIndex), GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex))
                    End If
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).type = TILE_TYPE_KEY Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    canMove = False
                    
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
                        canMove = False
                        
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
                        canMove = False
                        
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
            If map.Left > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            canMove = False
            Exit Function
        End If
    End If
        
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            ' Check to see if the map tile is blocked or not
            If map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).type = TILE_TYPE_BLOCKED Then
                canMove = False
                    If GetPlayerX(MyIndex) + 1 <= 15 Then
                        Call WarpTo_U(GetPlayerMap(MyIndex), GetPlayerMap(MyIndex), GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex))
                    End If
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).type = TILE_TYPE_KEY Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    canMove = False
                    
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
                        canMove = False
                        
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
                        canMove = False
                        
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
            If map.Right > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            canMove = False
            Exit Function
        End If
    End If
End Function

Sub CheckMovement()
    If GettingMap = False And canMoveNow = True Then
        If IsTryingToMove Then
            If canMove Then
                ' Check if player has the shift key down for running
                If ShiftDown Then
                    Player(MyIndex).moving = MOVING_RUNNING
                Else
                    Player(MyIndex).moving = MOVING_WALKING
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
                If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
                ' Gotta check :)
                If map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).type = TILE_TYPE_WARP_LEVEL Then
                    'PlaySound ("doorslam")
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
    frmMirage.clearPanes
    SaveMap = map
    InEditor = True
    frmMirage.picMapEditor.Visible = True
    DoEvents
    With frmMirage.picBack
        .Width = 22 * PIC_X
    End With
    With frmMirage.picBackSelect
        .Width = 22 * PIC_X
        '.Height = 50000 * PIC_Y
        .Picture = LoadPicture(App.Path + "\data\bmp\tiles0.bmp")
    End With
    With frmMirage
        .scrlPicture.Max = frmMirage.picBackSelect.Height / PIC_Y
        .opt1(0).value = True
        .Refresh
    End With
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single, TileSheet As Long)
Dim x1, y1 As Long

    If InEditor Then
        
   
    
        x1 = Int(x / PIC_X)
        y1 = Int(y / PIC_Y)
        
        
     'If frmMirage.optFringe = True Then
     If frmMirage.cmbLayers.ListIndex = 3 Then
     On Local Error Resume Next
        EditorTileSheet_Fringe = TileSheet
        EditorTileSheet_Anim = map.Tile(x1, y1).TileSheet_Anim
        EditorTileSheet_Mask = map.Tile(x1, y1).TileSheet_Mask
        EditorTileSheet_Ground = map.Tile(x1, y1).TileSheet_Ground
    'ElseIf frmMirage.optAnim = True Then
    ElseIf frmMirage.cmbLayers.ListIndex = 2 Then
    On Local Error Resume Next
        EditorTileSheet_Fringe = map.Tile(x1, y1).TileSheet_Fringe
        EditorTileSheet_Anim = TileSheet
        EditorTileSheet_Mask = map.Tile(x1, y1).TileSheet_Mask
        EditorTileSheet_Ground = map.Tile(x1, y1).TileSheet_Ground
    'ElseIf frmMirage.optMask = True Then
    ElseIf frmMirage.cmbLayers.ListIndex = 1 Then
    On Local Error Resume Next
        EditorTileSheet_Fringe = map.Tile(x1, y1).TileSheet_Fringe
        EditorTileSheet_Anim = map.Tile(x1, y1).TileSheet_Anim
        EditorTileSheet_Mask = TileSheet
        EditorTileSheet_Ground = map.Tile(x1, y1).TileSheet_Ground
    Else
    On Local Error Resume Next
        EditorTileSheet_Fringe = map.Tile(x1, y1).TileSheet_Fringe
        EditorTileSheet_Anim = map.Tile(x1, y1).TileSheet_Anim
        EditorTileSheet_Mask = map.Tile(x1, y1).TileSheet_Mask
        EditorTileSheet_Ground = TileSheet
    End If
        If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmMirage.optLayers.value = True Then
                With map.Tile(x1, y1)
                    If frmMirage.cmbLayers.ListIndex = 0 Then .Ground = EditorTileY * 50 + EditorTileX ' Ground Layer
                    If frmMirage.cmbLayers.ListIndex = 1 Then .mask = EditorTileY * 50 + EditorTileX ' Mask Layer
                    If frmMirage.cmbLayers.ListIndex = 2 Then .Anim = EditorTileY * 50 + EditorTileX ' Anim Layer
                    If frmMirage.cmbLayers.ListIndex = 3 Then .Fringe = EditorTileY * 50 + EditorTileX ' Fringe Layer
                    
                    'If frmMirage.optGround.Value = True Then .Ground = EditorTileY * 50 + EditorTileX
                    'If frmMirage.optMask.Value = True Then .Mask = EditorTileY * 50 + EditorTileX
                    'If frmMirage.optAnim.Value = True Then .Anim = EditorTileY * 50 + EditorTileX
                    'If frmMirage.optFringe.Value = True Then .Fringe = EditorTileY * 50 + EditorTileX
                    .TileSheet_Ground = EditorTileSheet_Ground
                    .TileSheet_Fringe = EditorTileSheet_Fringe
                    .TileSheet_Anim = EditorTileSheet_Anim
                    .TileSheet_Mask = EditorTileSheet_Mask
                End With
            Else
                With map.Tile(x1, y1)
                    If frmMirage.cmbAtributes.ListIndex = 0 Then .type = TILE_TYPE_BLOCKED
                    'If frmMirage.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                    'If frmMirage.optDamage.Value = True Then
                    If frmMirage.cmbAtributes.ListIndex = 7 Then
                    'LOOKY
                        .type = TILE_TYPE_DAMAGE
                        .Data1 = EditorDamage
                        .Data2 = 0
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
                        '.Data6 = 0
                        '.Data7 = 0
                        '.Data8 = 0
                        '.Data9 = 0
                        '.Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 8 Then
                    'If frmMirage.optHeal.Value = True Then
                    'LOOKY
                        .type = TILE_TYPE_HEAL
                        .Data1 = EditorDamage
                        .Data2 = 0
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 9 Then
                    'If frmMirage.optSign.Value = True Then
                    'LOOKY
                        .type = TILE_TYPE_SIGN
                        .Data1 = EditorSignNumber
                        .Data2 = 0
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 11 Then
                    'If frmMirage.optSign.Value = True Then
                    'LOOKY
                        .type = TILE_TYPE_LEVEL
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    
                    If frmMirage.cmbAtributes.ListIndex = 2 Then
                    'If frmMirage.optWarpDoor.Value = True Then
                    'LOOKY
                        .type = TILE_TYPE_WARP_LEVEL
                        .Data1 = EditorWarpMap
                        .Data2 = EditorWarpX
                        .Data3 = EditorWarpY
                        .Data4 = EditorMinLevel
                        .Data5 = EditorMaxLevel
'                        .Data6 = EditorMsg
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 1 Then
                    'If frmMirage.optWarp.Value = True Then
                        .type = TILE_TYPE_WARP
                        .Data1 = EditorWarpMap
                        .Data2 = EditorWarpX
                        .Data3 = EditorWarpY
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 3 Then
                    'If frmMirage.optItem.Value = True Then
                        .type = TILE_TYPE_ITEM
                        .Data1 = ItemEditorNum
                        .Data2 = ItemEditorValue
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 4 Then
                    'If frmMirage.optNpcAvoid.Value = True Then
                        .type = TILE_TYPE_NPCAVOID
                        .Data1 = 0
                        .Data2 = 0
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 5 Then
                    'If frmMirage.optKey.Value = True Then
                        .type = TILE_TYPE_KEY
                        .Data1 = KeyEditorNum
                        .Data2 = KeyEditorTake
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 6 Then
                    'If frmMirage.optKeyOpen.Value = True Then
                        .type = TILE_TYPE_KEYOPEN
                        .Data1 = KeyOpenEditorX
                        .Data2 = KeyOpenEditorY
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                    If frmMirage.cmbAtributes.ListIndex = 10 Then
                    'If frmMirage.optNpcSpawn.Value = True Then
                        .type = TILE_TYPE_NPC_SPAWN
                        .Data1 = EditorNPC_Num
                        .Data2 = 0
                        .Data3 = 0
                        .Data4 = 0
                        .Data5 = 0
'                        .Data6 = 0
'                        .Data7 = 0
'                        .Data8 = 0
'                        .Data9 = 0
'                        .Data10 = 0
                    End If
                End With
            End If
        End If
        
        If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmMirage.optLayers.value = True Then
                With map.Tile(x1, y1)
                    If frmMirage.cmbLayers.ListIndex = 0 Then .Ground = 0 ' Ground Layer
                    If frmMirage.cmbLayers.ListIndex = 1 Then .mask = 0 ' Mask Layer
                    If frmMirage.cmbLayers.ListIndex = 2 Then .Anim = 0 ' Anim Layer
                    If frmMirage.cmbLayers.ListIndex = 3 Then .Fringe = 0 ' Fringe Layer
                    
                    'If frmMirage.optGround.Value = True Then .Ground = 0
                    'If frmMirage.optMask.Value = True Then .Mask = 0
                    'If frmMirage.optAnim.Value = True Then .Anim = 0
                    'If frmMirage.optFringe.Value = True Then .Fringe = 0
                End With
            Else
                With map.Tile(x1, y1)
                    .type = 0
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .Data4 = 0
                    .Data5 = 0
'                    .Data6 = 0
'                    .Data7 = 0
'                    .Data8 = 0
'                    .Data9 = 0
'                    .Data10 = 0
                End With
            End If
        End If
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorTileX = Int(x / PIC_X)
        EditorTileY = Int(y / PIC_Y)
    End If
    Call BitBlt(frmMirage.picSelect.hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picBackSelect.hdc, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub EditorTileScroll()
    frmMirage.picBackSelect.top = (frmMirage.scrlPicture.value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    DoEvents
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    map = SaveMap
    InEditor = False
    frmMirage.picMapEditor.Visible = False
    frmMirage.txtSend.SetFocus
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, x As Long, y As Long

    ' Ground layer
    If frmMirage.cmbLayers.ListIndex = 0 Then ' Ground Layer
    'If frmMirage.optGround.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    map.Tile(x, y).Ground = 0
                Next x
            Next y
        End If
    End If

    ' Mask layer
    If frmMirage.cmbLayers.ListIndex = 1 Then ' Mask Layer
    'If frmMirage.optMask.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    map.Tile(x, y).mask = 0
                Next x
            Next y
        End If
    End If

    ' Animation layer
    If frmMirage.cmbLayers.ListIndex = 2 Then ' Anim Layer
    'If frmMirage.optAnim.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    map.Tile(x, y).Anim = 0
                Next x
            Next y
        End If
    End If

    ' Fringe layer
    If frmMirage.cmbLayers.ListIndex = 3 Then ' fringe Layer
    'If frmMirage.optFringe.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    map.Tile(x, y).Fringe = 0
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
                map.Tile(x, y).type = 0
            Next x
        Next y
    End If
End Sub

Public Sub ItemEditorInit()
On Error Resume Next
    
    frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\data\bmp\items.bmp")
    
    frmItemEditor.txtName.text = Trim(Item(EditorIndex).Name)
    frmItemEditor.scrlPic.value = Item(EditorIndex).Pic
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).type
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.scrlDurability.value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.value = Item(EditorIndex).Data2
        frmItemEditor.scrlBaseDamage.value = Item(EditorIndex).BaseDamage
        frmItemEditor.scrlStrength.value = Item(EditorIndex).str
        frmItemEditor.scrlIntel.value = Item(EditorIndex).intel
        frmItemEditor.scrlDex.value = Item(EditorIndex).dex
        frmItemEditor.scrlCon.value = Item(EditorIndex).con
        frmItemEditor.scrlWiz.value = Item(EditorIndex).wiz
        frmItemEditor.scrlCha.value = Item(EditorIndex).cha
        frmItemEditor.cmbWeaponType.ListIndex = Item(EditorIndex).weaponType
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Or (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_POTIONADDPP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Caption = "Spell Data"
    Else
        frmItemEditor.fraSpell.Caption = "Prayer Data"
    End If
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL Or frmItemEditor.cmbType.ListIndex = ITEM_TYPE_PRAYER) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraSpell.Visible = False
    End If
    frmItemEditor.txtDescription = Item(EditorIndex).Description
    
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).Name = frmItemEditor.txtName.text
    Item(EditorIndex).Pic = frmItemEditor.scrlPic.value
    Item(EditorIndex).type = frmItemEditor.cmbType.ListIndex

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.value
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).BaseDamage = frmItemEditor.scrlBaseDamage.value
        Item(EditorIndex).str = frmItemEditor.scrlStrength.value
        Item(EditorIndex).intel = frmItemEditor.scrlIntel.value
        Item(EditorIndex).dex = frmItemEditor.scrlDex.value
        Item(EditorIndex).con = frmItemEditor.scrlCon.value
        Item(EditorIndex).wiz = frmItemEditor.scrlWiz.value
        Item(EditorIndex).cha = frmItemEditor.scrlCha.value
        Item(EditorIndex).Poisons = frmItemEditor.chkPoison.value
        Item(EditorIndex).Poison_length = Val(frmItemEditor.txtPoisonLength)
        Item(EditorIndex).Poison_vital = Val(frmItemEditor.txtPoisonVital)
        Item(EditorIndex).weaponType = frmItemEditor.cmbWeaponType.ListIndex
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Or (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_POTIONADDPP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL Or frmItemEditor.cmbType.ListIndex = ITEM_TYPE_PRAYER) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    Item(EditorIndex).Description = frmItemEditor.txtDescription
    
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
    
End Sub

Public Sub ItemEditorBltItem()
    Call BitBlt(frmItemEditor.picPic.hdc, 0, 0, PIC_X, PIC_Y, frmItemEditor.picItems.hdc, 0, frmItemEditor.scrlPic.value * PIC_Y, SRCCOPY)
End Sub

Public Sub ItemLibBltItem(ByVal Number As Long)
    Call BitBlt(frmItemLib.picPic.hdc, 0, 0, PIC_X, PIC_Y, frmItemLib.picItems.hdc, 0, Number * PIC_Y, SRCCOPY)
End Sub

Public Sub MainBltInventItem(ByVal itemNumber As Long, ByVal slot As Long)
    Call BitBlt(frmMirage.picInvItem(slot).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, itemNumber * PIC_Y, SRCCOPY)
End Sub

Public Sub MainBltBankInvLocalItem(ByVal itemNumber As Long, ByVal slot As Long, ByVal startRow As Integer)
    Call BitBlt(frmMirage.picItemLocal(slot + (5 * startRow)).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, itemNumber * PIC_Y, SRCCOPY)
End Sub

Public Sub MainBltBankItem(ByVal itemNumber As Long, ByVal slot As Long)
    Call BitBlt(frmMirage.picbankItem(slot).hdc, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hdc, 0, itemNumber * PIC_Y, SRCCOPY)
End Sub

Public Sub CharGenBltSprite(ByVal Number As Long, ByVal count As Long)
    'Call BitBlt(frmNewChar.picChar.hdc, 0, 0, 32, 32, frmNewChar.picinit.hdc, 0, 0, SRCCOPY)
    
    Call StretchBlt(frmNewChar.picChar.hdc, 16, 16, PIC_X, PIC_Y, frmNewChar.picChars.hdc, count * PIC_X, _
        Number * PIC_Y, PIC_X, PIC_Y, SRCCOPY)
        
    'Call TransparentBlt(frmNewChar.picChar.hdc, 0, 0, PIC_X * 2, PIC_Y * 2, frmNewChar.picChars.hdc, count * PIC_X, _
        number * PIC_Y, PIC_X, PIC_Y, RGB(0, 0, 0))
        
End Sub


Public Sub NpcEditorInit()
On Error Resume Next
    
    frmNpcEditor.picsprites.Picture = LoadPicture(App.Path & "\data\bmp\sprites.bmp")
    
    frmNpcEditor.txtName.text = Trim(Npc(EditorIndex).Name)
    frmNpcEditor.txtAttackSay.text = Trim(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.value = Npc(EditorIndex).sprite
    frmNpcEditor.txtSpawnSecs.text = str(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.value = Npc(EditorIndex).Range
    frmNpcEditor.txtChance.text = str(Npc(EditorIndex).DropChance)
    frmNpcEditor.scrlNum.value = Npc(EditorIndex).DropItem
    frmNpcEditor.scrlValue.value = Npc(EditorIndex).DropItemValue
    frmNpcEditor.scrlSTR.value = Npc(EditorIndex).str
    frmNpcEditor.scrlDEF.value = Npc(EditorIndex).DEF
    frmNpcEditor.scrlSPEED.value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.value = Npc(EditorIndex).MAGI
    frmNpcEditor.txtExpGiven = Npc(EditorIndex).ExpGiven
    frmNpcEditor.lblExpGiven = Npc(EditorIndex).ExpGiven
    frmNpcEditor.txtStartHP = Npc(EditorIndex).HP
    frmNpcEditor.lblStartHP = Npc(EditorIndex).HP
    frmNpcEditor.optYesRespawn = Npc(EditorIndex).Respawn
    frmNpcEditor.txtQuestNo = Npc(EditorIndex).QuestID
    If frmNpcEditor.txtQuestNo > 0 Then frmNpcEditor.chkQuest = 1 Else frmNpcEditor.chkQuest = 0
    frmNpcEditor.optNoRespawn = Not frmNpcEditor.optYesRespawn
    If Npc(EditorIndex).Attack_with_Poison = True Then
        frmNpcEditor.chkPoisonAttack = 1
    Else
        frmNpcEditor.chkPoisonAttack = 0
    End If
    
    frmNpcEditor.cmdType.ListIndex = Npc(EditorIndex).type
    If Npc(EditorIndex).opensBank = True Then
        frmNpcEditor.chkBank = 1
    Else
        frmNpcEditor.chkBank = 0
    End If
    If Npc(EditorIndex).opensShop = True Then
        frmNpcEditor.chkShop = 1
    Else
        frmNpcEditor.chkShop = 0
    End If
    'frmNpcEditor.chkBank = Npc(EditorIndex).opensBank
    'frmNpcEditor.chkShop = Npc(EditorIndex).opensShop
    
    frmNpcEditor.txtPoisonLength = Npc(EditorIndex).Poison_length
    frmNpcEditor.txtPoisonVital = Npc(EditorIndex).Poison_vital
    
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).Name = frmNpcEditor.txtName.text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.text
    Npc(EditorIndex).sprite = frmNpcEditor.scrlSprite.value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.value
    Npc(EditorIndex).DropChance = Val(frmNpcEditor.txtChance.text)
    Npc(EditorIndex).DropItem = frmNpcEditor.scrlNum.value
    Npc(EditorIndex).DropItemValue = frmNpcEditor.scrlValue.value
    Npc(EditorIndex).str = frmNpcEditor.scrlSTR.value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.value
    Npc(EditorIndex).HP = frmNpcEditor.txtStartHP
    Npc(EditorIndex).ExpGiven = frmNpcEditor.txtExpGiven
    Npc(EditorIndex).Respawn = frmNpcEditor.optYesRespawn.value
    Npc(EditorIndex).Attack_with_Poison = CBool(frmNpcEditor.chkPoisonAttack.value)
    Npc(EditorIndex).Poison_length = CLng(frmNpcEditor.txtPoisonLength)
    Npc(EditorIndex).Poison_vital = CLng(frmNpcEditor.txtPoisonVital)
    Npc(EditorIndex).opensBank = CBool(frmNpcEditor.chkBank)
    Npc(EditorIndex).opensShop = CBool(frmNpcEditor.chkShop)
    Npc(EditorIndex).type = frmNpcEditor.cmdType.ListIndex
    
    If frmNpcEditor.chkQuest Then
        Npc(EditorIndex).QuestID = frmNpcEditor.txtQuestNo
    Else
        Npc(EditorIndex).QuestID = 0
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
    Call BitBlt(frmNpcEditor.picsprite.hdc, 0, 0, PIC_X, PIC_Y, frmNpcEditor.picsprites.hdc, 3 * PIC_X, frmNpcEditor.scrlSprite.value * PIC_Y, SRCCOPY)
End Sub

Public Sub ChooseBltSprite()
    Call BitBlt(frmSprite.picsprite.hdc, 0, 0, PIC_X, PIC_Y, frmSprite.picsprites.hdc, 3 * PIC_X, frmSprite.scrlSprite.value * PIC_Y, SRCCOPY)
End Sub

Public Sub ShopEditorInit()
On Error Resume Next

Dim i As Long

    frmShopEditor.txtName.text = Trim(Shop(EditorIndex).Name)
    frmShopEditor.txtJoinSay.text = Trim(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.text = Trim(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.value = Shop(EditorIndex).FixesItems
    
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
    Shop(EditorIndex).Name = frmShopEditor.txtName.text
    Shop(EditorIndex).JoinSay = frmShopEditor.txtJoinSay.text
    Shop(EditorIndex).LeaveSay = frmShopEditor.txtLeaveSay.text
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
On Error Resume Next

Dim i As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim(Class(i).Name)
    Next i
    
    frmSpellEditor.txtName.text = Trim(Spell(EditorIndex).Name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.value = Spell(EditorIndex).LevelReq
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).type
    If Spell(EditorIndex).type <> SPELL_TYPE_GIVEITEM Then
        frmSpellEditor.fraVitals.Visible = True
        frmSpellEditor.fraGiveItem.Visible = False
        frmSpellEditor.scrlVitalMod.value = Spell(EditorIndex).Data1
    Else
        frmSpellEditor.fraVitals.Visible = False
        frmSpellEditor.fraGiveItem.Visible = True
        frmSpellEditor.scrlItemNum.value = Spell(EditorIndex).Data1
        frmSpellEditor.scrlItemValue.value = Spell(EditorIndex).Data2
    End If
    frmSpellEditor.scrlMusic = Spell(EditorIndex).Sound
    
        
    frmSpellEditor.Show vbModal
End Sub

Public Sub PrayerEditorInit()
On Error Resume Next

Dim i As Long

    frmPrayerEditor.cmbClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmPrayerEditor.cmbClassReq.AddItem Trim(Class(i).Name)
    Next i
    
    frmPrayerEditor.txtName.text = Trim(Prayer(EditorIndex).Name)
    frmPrayerEditor.cmbClassReq.ListIndex = Prayer(EditorIndex).ClassReq
    frmPrayerEditor.scrlLevelReq.value = Prayer(EditorIndex).LevelReq
    frmPrayerEditor.cmbType.ListIndex = Prayer(EditorIndex).type
        
'    frmPrayerEditor.cmbType.ListIndex = Spell(EditorIndex).Type
'    If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
'        frmSpellEditor.fraVitals.Visible = True
'        frmSpellEditor.fraGiveItem.Visible = False
'        frmSpellEditor.scrlVitalMod.value = Spell(EditorIndex).Data1
'    Else
'        frmSpellEditor.fraVitals.Visible = False
'        frmSpellEditor.fraGiveItem.Visible = True
'        frmSpellEditor.scrlItemNum.value = Spell(EditorIndex).Data1
'        frmSpellEditor.scrlItemValue.value = Spell(EditorIndex).Data2
'    End If
'    frmSpellEditor.scrlMusic = Spell(EditorIndex).sound
    
        
    frmPrayerEditor.Show vbModal
End Sub

Public Sub QuestEditorInit()
On Error Resume Next

    frmQuestEditor.txtQuestFinishText.text = Quests(EditorIndex).FinishQuestMessage
    frmQuestEditor.txtQuestStartText.text = Quests(EditorIndex).StartQuestMsg
    frmQuestEditor.txtQuestMiddleText.text = Quests(EditorIndex).GetItemQuestMsg
    frmQuestEditor.scrItemCol = Quests(EditorIndex).ItemToObtain
    frmQuestEditor.txtItemDesc = Item(Quests(EditorIndex).ItemToObtain).Name
    frmQuestEditor.txtItemNo = Quests(EditorIndex).ItemToObtain
    frmQuestEditor.scrItemGiv = Quests(EditorIndex).ItemGiven
    frmQuestEditor.txtItemDesc2 = Item(Quests(EditorIndex).ItemGiven).Name
    frmQuestEditor.txtItemNum2 = Quests(EditorIndex).ItemGiven
    frmQuestEditor.txtItemVal2 = Quests(EditorIndex).ItemValGiven
    frmQuestEditor.txtexp = Quests(EditorIndex).ExpGiven
    frmQuestEditor.txtMinLevel = Quests(EditorIndex).requiredLevel
    frmQuestEditor.txtGoldGiven = Quests(EditorIndex).goldGiven

    frmQuestEditor.Show vbModal
End Sub

Public Sub SignEditorInit()
'On Error Resume Next

Dim i As Long
    frmSignEditor.txtheader.text = Trim(Signs(EditorIndex).header)
    frmSignEditor.txtmsg.text = Trim(Signs(EditorIndex).Msg)
    frmSignEditor.Show vbModal
End Sub
Public Sub SignEditorOk()
    Signs(EditorIndex).header = frmSignEditor.txtheader.text
    Signs(EditorIndex).Msg = frmSignEditor.txtmsg.text
    Call SendSaveSign(EditorIndex)
    InSignEditor = False
    Unload frmSignEditor
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).Name = frmSpellEditor.txtName.text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.value
    Spell(EditorIndex).type = frmSpellEditor.cmbType.ListIndex
    If Spell(EditorIndex).type <> SPELL_TYPE_GIVEITEM Then
        Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.value
    Else
        Spell(EditorIndex).Data1 = frmSpellEditor.scrlItemNum.value
        Spell(EditorIndex).Data2 = frmSpellEditor.scrlItemValue.value
    End If
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).Sound = frmSpellEditor.scrlMusic
    Spell(EditorIndex).ManaUse = frmSpellEditor.txtMana
    
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub PrayerEditorOk()
    Prayer(EditorIndex).Name = frmPrayerEditor.txtName.text
    Prayer(EditorIndex).ClassReq = frmPrayerEditor.cmbClassReq.ListIndex
    Prayer(EditorIndex).LevelReq = frmPrayerEditor.scrlLevelReq.value
    Prayer(EditorIndex).type = frmPrayerEditor.cmbType.ListIndex
    Prayer(EditorIndex).Data1 = frmPrayerEditor.scrlVitalMod.value
    Prayer(EditorIndex).ManaUse = frmPrayerEditor.txtMana
    
    Call SendSavePrayer(EditorIndex)
    InPrayerEditor = False
    Unload frmPrayerEditor
End Sub


Public Sub QuestEditorOk()
    Quests(EditorIndex).ExpGiven = Val(frmQuestEditor.txtexp)
    Quests(EditorIndex).FinishQuestMessage = frmQuestEditor.txtQuestFinishText
    Quests(EditorIndex).StartQuestMsg = frmQuestEditor.txtQuestStartText
    Quests(EditorIndex).GetItemQuestMsg = frmQuestEditor.txtQuestMiddleText
    Quests(EditorIndex).ItemToObtain = Val(frmQuestEditor.txtItemNo)
    Quests(EditorIndex).ItemGiven = Val(frmQuestEditor.txtItemNum2)
    Quests(EditorIndex).ItemValGiven = Val(frmQuestEditor.txtItemVal2)
    Quests(EditorIndex).requiredLevel = Val(frmQuestEditor.txtMinLevel)
    Quests(EditorIndex).goldGiven = Val(frmQuestEditor.txtGoldGiven)
    
    Call SendSaveQuest(EditorIndex)
    InquestEditor = False
    Unload frmQuestEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub
Public Sub PrayerEditorCancel()
    InPrayerEditor = False
    Unload frmPrayerEditor
End Sub

Public Sub questEditorCancel()
    InquestEditor = False
    Unload frmQuestEditor
End Sub

Public Sub UpdateInventory()
Dim i As Long

    frmMirage.lstInv.Clear
    For i = 0 To frmMirage.picInvItem.UBound - 1
        frmMirage.picInvItem(i).BackColor = vbBlack
    Next i
    'Debug.Print (GetPlayerWeaponSlot(MyIndex))
    'Debug.Print (GetPlayerArmorSlot(MyIndex))
    'Debug.Print (GetPlayerHelmetSlot(MyIndex))
    'Debug.Print (GetPlayerShieldSlot(MyIndex))
    If GetPlayerWeaponSlot(MyIndex) = (selectedInvItem + 1) Or GetPlayerArmorSlot(MyIndex) = (selectedInvItem + 1) Or GetPlayerHelmetSlot(MyIndex) = (selectedInvItem + 1) Or GetPlayerShieldSlot(MyIndex) = (selectedInvItem + 1) Then
        frmMirage.picInvItem(selectedInvItem).BackColor = QBColor(12)
    Else
        frmMirage.picInvItem(selectedInvItem).BackColor = QBColor(14)
    End If
    'frmMirage.picInvItem(selectedInvItem).BackColor = QBColor(14)
    ' Show the inventory
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
         frmMirage.picInvItem(i - 1).ToolTipText = Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " - " & GetPlayerInvItemValue(MyIndex, i)
         Call MainBltInventItem(Item(GetPlayerInvItemNum(MyIndex, i)).Pic, i - 1)
            If Item(GetPlayerInvItemNum(MyIndex, i)).type = ITEM_TYPE_CURRENCY Then
                frmMirage.lstInv.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Then
                    frmMirage.lstInv.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                    frmMirage.picInvItem(i - 1).ToolTipText = frmMirage.picInvItem(i - 1).ToolTipText & " (worn)"
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

Public Sub UpdateBank(Optional ByVal startRow As Integer = 0)
frmMirage.picBankWindow.Visible = True
InBank = True
frmMirage.lblInvGold.Caption = 0
frmMirage.lblBankGold.Caption = 0
If startRow > 4 Then startRow = 4
If startRow < 0 Then startRow = 0
Dim i As Long
For i = 0 To 29 Step 1
    frmMirage.picbankItem(i).BackColor = vbBlack
    frmMirage.picItemLocal(i).BackColor = vbBlack
    frmMirage.picbankItem(selectedBankItem).BackColor = QBColor(14)
    frmMirage.picItemLocal(selectedBankInvLocalItem).BackColor = QBColor(14)
Next i
    For i = 1 To 30 Step 1
        If (GetPlayerInvItemNum(MyIndex, i) > 0) And (GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS) Then
            If GetPlayerInvItemNum(MyIndex, i) = 2 Then
                frmMirage.lblInvGold.Caption = CLng(frmMirage.lblInvGold.Caption) + GetPlayerInvItemValue(MyIndex, i)
                'Call MainBltBankInvLocalItem(Item(GetPlayerInvItemNum(MyIndex, i)).Pic, i - 1, startRow)
            Else
                Call MainBltBankInvLocalItem(Item(GetPlayerInvItemNum(MyIndex, i)).Pic, i - 1, startRow)
                'Debug.Print "inv num: " & GetPlayerInvItemNum(MyIndex, i)
            End If
        End If
            'On Local Error Resume Next
            'Debug.Print "thing: " & GetPlayerBankItemNum(MyIndex, i)
        If GetPlayerBankItemNum(MyIndex, i) > 0 And GetPlayerBankItemNum(MyIndex, i) < MAX_BANK Then
            If GetPlayerBankItemNum(MyIndex, i) = 2 Then
                frmMirage.lblBankGold.Caption = CLng(frmMirage.lblBankGold.Caption) + GetPlayerBankItemValue(MyIndex, i)
                'Call MainBltBankItem(Item(GetPlayerBankItemNum(MyIndex, i)).Pic, i - 1)
            Else
                'Debug.Print "thing: " & GetPlayerBankItemNum(MyIndex, i)
                Call MainBltBankItem(Item(GetPlayerBankItemNum(MyIndex, i)).Pic, i - 1)
            End If
        End If
    Next i
    
    
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1 As Long, y1 As Long

    x1 = Int(x / PIC_X)
    y1 = Int(y / PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
    If Button > 1 Then
        Call SendData("getCords" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
    If Button > 2 Then
        Call WarpTo(GetPlayerMap(MyIndex), GetPlayerMap(MyIndex), x1, y1)
    End If
End Sub



Public Sub SquareAlphaBlend( _
ByVal cSrc_Widht As Integer, _
ByVal cSrc_Height As Integer, _
ByVal cSrc As Long, _
ByVal cSrc_X As Integer, _
ByVal cSrc_Y As Integer, _
ByVal cDest As Long, _
ByVal cDest_X As Integer, _
ByVal cDest_Y As Integer, _
ByVal nLevel As Byte)

    Dim LrProps As rBlendProps
    Dim LnBlendPtr As Long
    
    LrProps.tBlendAmount = nLevel
    CopyMemory LnBlendPtr, LrProps, 4
    
    AlphaBlend cDest, cDest_X, cDest_Y, cSrc_Widht, cSrc_Height, cSrc, cSrc_X, cSrc_Y, cSrc_Widht, cSrc_Height, LnBlendPtr
    
End Sub

