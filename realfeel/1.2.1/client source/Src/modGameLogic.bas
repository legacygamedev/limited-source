Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub Main()
Dim i As Long, f As Integer, FilePath As String
f = FreeFile
        
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    
    ' Clear out players
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
    Next i
    Call ClearTempTile
    
    FilePath = App.Path & "/data.dat"
If Not FileExist("data.dat") Then
    Open FilePath For Output As #f
        Print #f, ";Do not delete this file!"
        Print #f, ""
        Print #f, "[Account]"
        Print #f, "Name="
        Print #f, "Password="
        Print #f, ""
        Print #f, "[Address]"
        Print #f, "IP="
        Print #f, "Port="
        Print #f, ""
        Print #f, "[Settings]"
        Print #f, "StrokeText=1"
        Print #f, "Music=1"
        Print #f, "Sound=1"
        Print #f, "NPC_Names=1"
        Print #f, ""
        Print #f, "[FilePaths]"
        Print #f, "DLLs= \DLLs\"
        Print #f, "GFX= \Gfx\"
        Print #f, "LOGS= \logs\"
        Print #f, "MAPS= \maps\"
        Print #f, "MUSIC = \music\"
        Print #f, "SOUND = \sound\"
        Print #f, "GUI = \Gfx\GUI\"
        Print #f, ""
        Print #f, ";Setting the style to FIXED chooses the target music to play"
        Print #f, ";Setting the style to RANDOM chooses a piece of music from the pointer if the pointer is a folder"
        Print #f, ";Setting the style to OFF quits the music player"
        Print #f, ";Setting the playmode to NORMAL runs the song at normal speed"
        Print #f, ";Setting the playmode to SLOW runs the song at a slower speed"
        Print #f, ";Setting the playmode to FAST runs the song at a faster speed"
        Print #f, ";Setting the intro loop to TRUE loops the music"
        Print #f, ";Setting the intro loop to FALSE does not loop the music"
        Print #f, "[IntroMusic]"
        Print #f, "INTRO_POINTER = \music\"
        Print #f, "INTRO_STYLE = FIXED"
        Print #f, "INTRO_PLAYMODE = NORMAL"
        Print #f, "INTRO_PLAYLOOP = FALSE"
    Close #f
End If

    ' Check if a DLL or OCX file is missing, and bring up the register menu
    If CheckReg = True Then
        Call Main2
    Else
        frmRegFiles.Visible = True
    End If
End Sub

Sub Main2()
    frmSendGetData.Visible = True
    Call SetStatus("Initializing TCP settings...")
    DoEvents
    Call TcpInit
    Call SetStatus("Booting Interphase Core Files...")
    DoEvents
    Call DataCheck
    Call SetStatus("Initializing DirectX8...")
    DoEvents
    While Not InitDirectX8
    Wend
    Call SetStatus("Updating information...")
    DoEvents
    Call LoadMenu
    frmSendGetData.Visible = False
End Sub

Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
End Sub

Public Sub MenuState(ByVal State As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added website constant.
'****************************************************************
    
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
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMainMenu.txtName.Text, frmMainMenu.txtPassword.Text)
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

    If Not IsConnected Then
        Call LoadMenu
        frmSendGetData.Visible = False
        Call MsgBox("Error connecting to the specified server.  Please try to reconnect in a few minutes or visit " & WEBSITE & " for technical help.", vbOKOnly, GAME_NAME)
    End If
End Sub

Public Sub GameLoop()
On Error GoTo ErrorHandler:

Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim X As Long
Dim Y As Long
Dim i As Long
Dim rec_back As RECT

' FPS Cap data
Dim FPS_CAP As Long
Dim FPS_MAX_COUNT As Long
Dim FPS_SINGLE_COUNT As Long
Dim FPS_COUNT As Long
Dim FPS_MILLI As Long

    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0
    
    ' Set FPS data
    FPS_CAP = 30
    FPS_COUNT = 0
    FPS_MAX_COUNT = 0
    FPS_SINGLE_COUNT = 0

    Do While InGame
    ' Useless tick count
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
    
        If frmDualSolace.WindowState <> 1 Then
        Call Direct3D.ClearScreen(RGB(0, 0, 0))
        Call Direct3D.BeginScene
        Call Direct3D.BeginDrawing
        If PauseMap = False Then
            
            ' Blit out tiles layers ground/anim1/anim2
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Call BltTile(X, Y)
                Next X
            Next Y
                    
            ' Blit out the items
            For i = 1 To MAX_MAP_ITEMS
                If MapItem(i).num > 0 Then
                    Call BltItem(i)
                End If
            Next i
        
            ' Blit out the npcs
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Death.Display = False Then
                    Call BltNpc(i)
                ElseIf MapNpc(i).Death.Display = True Then
                    Call BltNpcDeath(i)
                End If
            Next i
        
            ' Blit out players
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) And Player(i).Death.Display = False Then
                        Call BltPlayer(i)
                    ElseIf Player(i).Death.Display = True Then
                        If Player(i).Death.DeathMap = GetPlayerMap(MyIndex) Then
                            Call BltPlayerDeath(i)
                        End If
                    End If
                End If
            Next i
                
            ' Blit out tile layer fringe
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Call BltFringeTile(X, Y)
                Next X
            Next Y
        
            'Draw Player names and damage
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    Call BltPlayerName(i)
                    Call DrawDamage(i, "PLAYER")
                End If
            Next i
        
            'See if the client should blt the npc names
            'Draw the damage either way
            If frmGameSettings.chkNPCNames.Value = 1 Then
                For i = 1 To MAX_MAP_NPCS
                    If MapNpc(i).num > 0 Then
                        Call BltNpcName(i)
                        Call DrawDamage(i, "NPC")
                    End If
                Next i
            Else
                For i = 1 To MAX_MAP_NPCS
                    If MapNpc(i).num > 0 Then
                        Call DrawDamage(i, "NPC")
                    End If
                Next i
            End If
        
            ' Stroke Map name
            'Call Direct3D.DrawText(Trim$(Map.Name), C_Black, DT_LEFT, (frmDualSolace.picScreen.ScaleWidth \ 2) - (Len(Trim$(Map.Name)) * 10) + 1, 1)
            'Call Direct3D.DrawText(Trim$(Map.Name), C_Black, DT_LEFT, (frmDualSolace.picScreen.ScaleWidth \ 2) - (Len(Trim$(Map.Name)) * 10) - 1, 1)
            'Call Direct3D.DrawText(Trim$(Map.Name), C_Black, DT_LEFT, (frmDualSolace.picScreen.ScaleWidth \ 2) - (Len(Trim$(Map.Name)) * 10), 1 + 1)
            'Call Direct3D.DrawText(Trim$(Map.Name), C_Black, DT_LEFT, (frmDualSolace.picScreen.ScaleWidth \ 2) - (Len(Trim$(Map.Name)) * 10), 1 - 1)
        
            ' Draw map name
            If Map.Moral = MAP_MORAL_NONE Then
                ' Reset the font settings - add bold
                Call Direct3D.SetupFont("Verdana", "12", True, False)
                
                ' Draw the text
                Call Direct3D.DrawText(Trim$(Map.Name), C_Red, DT_LEFT, (frmDualSolace.picScreen.ScaleWidth \ 2) - (Len(Trim$(Map.Name)) * 10), 1)
            Else
                ' Reset the font settings - add bold
                Call Direct3D.SetupFont("Verdana", "12", True, False)
                
                ' Draw the text
                Call Direct3D.DrawText(Trim$(Map.Name), C_White, DT_LEFT, (frmDualSolace.picScreen.ScaleWidth \ 2) - (Len(Trim$(Map.Name)) * 10), 1)
            End If
        
            ' Check if we are getting a map, and if we are tell them so
            If GettingMap = True Then
                ' Reset the font settings
                Call Direct3D.SetupFont("Verdana", 12, False, False)
                
                ' Draw the text
                Call Direct3D.DrawText("Receiving Map...", C_BrightCyan, DT_LEFT, 1, 1)
            End If
            
            'Draw an A for attributes if any exist besides walkable
            If InEditor = True Then
                If frmDualSolace.optAttribs.Value = True Then
                    If DepictAttributeTiles = True Then
                        For Y = 0 To MAX_MAPY
                            For X = 0 To MAX_MAPX
                                With Map.Tile(X, Y)
                                    ' Reset the font settings
                                    Call Direct3D.SetupFont("Verdana", 12, False, False)
                                    
                                    ' Draw the text
                                    If .Walkable = False Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .Blocked = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .North = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .West = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .East = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .South = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .Warp = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .Item = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .NpcAvoid = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .Key = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .KeyOpen = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .Bank = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .Shop = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .Heal = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                    If .Damage = True Then Call Direct3D.DrawText("A", C_White, DT_LEFT, X * PIC_X + 8, Y * PIC_Y + 8)
                                End With
                            Next X
                        Next Y
                    End If
                End If
            End If
            
            'Check to see if in map editor, if so, draw the mapeditor texture
            If InEditor = True Then
                If frmDualSolace.optAttribs.Value = True Then
                    If AttributeDisplay = True Then
                        Call DrawMapEditorTexture(Mouse_X, Mouse_Y)
                    End If
                End If
            End If
            
            ' Reset the font settings
            Call Direct3D.SetupFont("Verdana", 12, False, False)
            
            'Draw the FPS in the bottom left
            Call Direct3D.DrawText(CStr(GameFPS), C_White, DT_LEFT, 1, frmDualSolace.picScreen.ScaleHeight - 20)
        ElseIf PauseMap = True Then
            ' Reset the font settings
            Call Direct3D.SetupFont("Verdana", 12, False, False)
            
            ' Draw the text
            If PauseMessage <> "" Then Call Direct3D.DrawText(PauseMessage, C_White, DT_LEFT, Int(frmDualSolace.picScreen.ScaleWidth \ 2) - (10 * Len(PauseMessage)), Int(frmDualSolace.picScreen.ScaleHeight \ 2))
        End If
        Call Direct3D.EndDrawing
        Call Direct3D.EndScene
            ' Blit the backbuffer
            Call Direct3D.RenderScreen
        End If
        
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
            
            ' If the player is using the map editor, and has his
            ' player locked in, we need to prevent this slowdown
            ' The second condition is added because the player moves
            ' like crazy, not good for the client
            
            'If InEditor = False Then
            '    Sleep 13
            '    DoEvents
            'ElseIf InEditor = True And AllowMovement = True Then
            '    Sleep 13
            '    DoEvents
            'End If
            
            
            
            ' New FPS lock system
            'If GameFPS > FPS_CAP Then
            '    If FPS_COUNT <= 0 Then
                    ' Determine the number of milliseconds per frame
            '        FPS_MAX_COUNT = Round((1000 - ((1000 \ (GameFPS)) * (FPS_CAP))), 0)
            '        'FPS_SINGLE_COUNT = Round((FPS_MAX_COUNT \ (FPS_CAP)), 0)
            '        FPS_SINGLE_COUNT = Round((FPS_MAX_COUNT \ (GameFPS - FPS_CAP)), 0)
            '    End If
            'End If
            
            ' Execute slowdown for FPS if the lock has been set
            'If FPS_MAX_COUNT <> 0 Then
            '   FPS_COUNT = GetTickCount
            '   Sleep FPS_SINGLE_COUNT
            '   FPS_MAX_COUNT = FPS_MAX_COUNT - FPS_SINGLE_COUNT
            
            '    If FPS_MAX_COUNT < FPS_SINGLE_COUNT Then
            '        FPS_COUNT = 0
            '    End If
            'End If
            
            ' Returned old FPS Lock, newer version caused random slowdowns.
            Do While GetTickCount < Tick + 32
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
    
    frmDualSolace.Visible = False
    frmSendGetData.Visible = True
    Call SetStatus("Destroying game data...")
    
    ' Shutdown the game
    Call GameDestroy
    
    ' Report disconnection if server disconnects
    If IsConnected = False Then
        Call MsgBox("Thank you for playing " & GAME_NAME & "!", vbOKOnly, GAME_NAME)
    End If
    
    Exit Sub
    
ErrorHandler:
MsgBox "Error #: " & Err.Number & vbCrLf & "Error Description: " & Err.Description
End Sub

Public Sub GameDestroy()
    Call DestroyDirectX8
    End
End Sub

Public Sub DrawDamage(ByVal num As Long, ByVal Focus As String)
Dim n As Byte
Select Case LCase$(Trim$(Focus))

Case "npc":
    ' Loop through, update and draw any damages
    For n = 1 To 5
        If MapNpc(num).Damage(n).Display = True Then
            ' Add one to the animation count
            MapNpc(num).Damage(n).AnimCount = MapNpc(num).Damage(n).AnimCount + 1
        
            ' Draw the text
            Call Direct3D.DrawText(CStr(MapNpc(num).Damage(n).Value), C_BrightRed, DT_LEFT, MapNpc(num).Damage(n).TextX, MapNpc(num).Damage(n).TextY - MapNpc(num).Damage(n).AnimCount)
            
            ' Check if the damage drawing has passed through 5 cycles
            If MapNpc(num).Damage(n).AnimCount > 5 Then
                ' Reset all data
                MapNpc(num).Damage(n).Display = False
                MapNpc(num).Damage(n).AnimCount = 0
                MapNpc(num).Damage(n).TextX = 0
                MapNpc(num).Damage(n).TextY = 0
                MapNpc(num).Damage(n).Value = 0
            End If
        End If
    Next n
    
Case "player":
    ' Loop through, update and draw any damages
    For n = 1 To 5
        If Player(num).Damage(n).Display = True Then
            ' Add one to the animation count
            Player(num).Damage(n).AnimCount = Player(num).Damage(n).AnimCount + 1
        
            ' Draw the text
            Call Direct3D.DrawText(CStr(Player(num).Damage(n).Value), C_BrightRed, DT_LEFT, Player(num).Damage(n).TextX, Player(num).Damage(n).TextY - Player(num).Damage(n).AnimCount)
            
            ' Check if the damage drawing has passed through 5 cycles
            If Player(num).Damage(n).AnimCount > 5 Then
                ' Reset all data
                Player(num).Damage(n).Display = False
                Player(num).Damage(n).AnimCount = 0
                Player(num).Damage(n).TextX = 0
                Player(num).Damage(n).TextY = 0
                Player(num).Damage(n).Value = 0
            End If
        End If
    Next n
    
End Select
End Sub

Public Sub DrawMapEditorTexture(ByVal X As Single, ByVal Y As Single)
Dim MoveX As Long, MoveY As Long
Dim DrawOffsetX As Long, DrawOffsetY As Long
Dim PicWidth As Long, PicHeight As Long
Dim PosX As Long, PosY As Long
Dim rec As DxVBLibA.RECT 'Added for Direct3D -smchronos
Dim Vect_Position As D3DVECTOR2 'Added for Direct3D -smchronos

PicWidth = Direct3D.GetPicWidth(App.Path + GFX_PATH + "attribute_display" + GFX_EXT)
PicHeight = Direct3D.GetPicHeight(App.Path + GFX_PATH + "attribute_display" + GFX_EXT)

'Set the offsets
DrawOffsetX = 5
DrawOffsetY = 5

If X + PicWidth + DrawOffsetX > frmDualSolace.picScreen.ScaleWidth Then
    MoveX = frmDualSolace.picScreen.ScaleWidth - (X + PicWidth + DrawOffsetX)
    PosX = X + MoveX
Else
    PosX = X + DrawOffsetX
End If
If Y + PicHeight + DrawOffsetY > frmDualSolace.picScreen.ScaleHeight Then
    MoveY = frmDualSolace.picScreen.ScaleHeight - (Y + PicHeight + DrawOffsetY)
    PosY = Y + MoveY
Else
    PosY = Y + DrawOffsetY
End If

Vect_Position.X = PosX
Vect_Position.Y = PosY

With rec
    rec.top = 1
    rec.Left = 1
    rec.Right = PicWidth
    rec.bottom = PicHeight
End With

Call Direct3D.DrawTex(mapeditortex, rec, Vect_Position, 1, &HAAAAAAAA)

' Reset the font settings
Call Direct3D.SetupFont("Verdana", 12, False, False)

' Draw the text
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Walkable), C_White, DT_LEFT, PosX + 105, PosY + 37)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Blocked), C_White, DT_LEFT, PosX + 91, PosY + 74)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).North), C_White, DT_LEFT, PosX + 121, PosY + 96)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).South), C_White, DT_LEFT, PosX + 123, PosY + 113)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).West), C_White, DT_LEFT, PosX + 119, PosY + 131)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).East), C_White, DT_LEFT, PosX + 113, PosY + 149)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Warp), C_White, DT_LEFT, PosX + 66, PosY + 185)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).WarpMap), C_White, DT_LEFT, PosX + 109, PosY + 202)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).WarpX), C_White, DT_LEFT, PosX + 88, PosY + 221)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).WarpY), C_White, DT_LEFT, PosX + 88, PosY + 239)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Item), C_White, DT_LEFT, PosX + 62, PosY + 275)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).ItemNum), C_White, DT_LEFT, PosX + 110, PosY + 292)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).ItemValue), C_White, DT_LEFT, PosX + 120, PosY + 311)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).NpcAvoid), C_White, DT_LEFT, PosX + 113, PosY + 346)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Key), C_White, DT_LEFT, PosX + 55, PosY + 382)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).KeyNum), C_White, DT_LEFT, PosX + 101, PosY + 400)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).KeyTake), C_White, DT_LEFT, PosX + 105, PosY + 418)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).KeyOpen), C_White, DT_LEFT, PosX + 276, PosY + 36)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).KeyOpenX), C_White, DT_LEFT, PosX + 298, PosY + 53)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).KeyOpenY), C_White, DT_LEFT, PosX + 298, PosY + 72)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Shop), C_White, DT_LEFT, PosX + 244, PosY + 109)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).ShopNum), C_White, DT_LEFT, PosX + 291, PosY + 126)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Bank), C_White, DT_LEFT, PosX + 245, PosY + 162)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Heal), C_White, DT_LEFT, PosX + 242, PosY + 197)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).HealValue), C_White, DT_LEFT, PosX + 295, PosY + 216)
Call Direct3D.DrawText(CBool(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).Damage), C_White, DT_LEFT, PosX + 270, PosY + 253)
Call Direct3D.DrawText(CStr(Map.Tile(Int(X \ PIC_X), Int(Y \ PIC_Y)).DamageValue), C_White, DT_LEFT, PosX + 298, PosY + 270)
End Sub

Public Sub BltTile(ByVal X As Long, ByVal Y As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Ground As Long
Dim Anim1 As Long
Dim Anim2 As Long
Dim Anim3, Anim4 As Long
Dim rec As DxVBLibA.RECT 'Added for Direct3D -smchronos
Dim Vect_Position As D3DVECTOR2 'Added for Direct3D -smchronos

Vect_Position.X = X * PIC_X
Vect_Position.Y = Y * PIC_Y

    With Map.Tile(X, Y)
        Ground = .Ground
        Anim1 = .Mask
        Anim2 = .Anim
        Anim3 = .Mask2
        Anim4 = .Anim2
    End With
    
    With rec
        .top = Int(Ground / 7) * PIC_Y
        .bottom = .top + PIC_Y
        .Left = (Ground - Int(Ground / 7) * 7) * PIC_X
        .Right = .Left + PIC_X
    End With
    Call Direct3D.DrawTex(tiletex, rec, Vect_Position)
    
    If (MapAnim = 0) Or (Anim2 <= 0) Then
        ' Is there an animation tile to plot?
        If Anim1 > 0 And TempTile(X, Y).DoorOpen = NO Then
            With rec
                .top = Int(Anim1 / 7) * PIC_Y
                .bottom = .top + PIC_Y
                .Left = (Anim1 - Int(Anim1 / 7) * 7) * PIC_X
                .Right = .Left + PIC_X
            End With
            Call Direct3D.DrawTex(tiletex, rec, Vect_Position)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim2 > 0 Then
            With rec
                .top = Int(Anim2 / 7) * PIC_Y
                .bottom = .top + PIC_Y
                .Left = (Anim2 - Int(Anim2 / 7) * 7) * PIC_X
                .Right = .Left + PIC_X
            End With
            Call Direct3D.DrawTex(tiletex, rec, Vect_Position)
        End If
    End If
    
    If (MapAnim = 0) Or (Anim4 <= 0) Then
        ' Is there an animation tile to plot?
        If Anim3 > 0 And TempTile(X, Y).DoorOpen = NO Then
            With rec
                .top = Int(Anim3 / 7) * PIC_Y
                .bottom = .top + PIC_Y
                .Left = (Anim3 - Int(Anim3 / 7) * 7) * PIC_X
                .Right = .Left + PIC_X
            End With
            Call Direct3D.DrawTex(tiletex, rec, Vect_Position)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim4 > 0 Then
            With rec
                .top = Int(Anim4 / 7) * PIC_Y
                .bottom = .top + PIC_Y
                .Left = (Anim4 - Int(Anim4 / 7) * 7) * PIC_X
                .Right = .Left + PIC_X
            End With
            Call Direct3D.DrawTex(tiletex, rec, Vect_Position)
        End If
    End If
End Sub

Public Sub BltItem(ByVal ItemNum As Long)
'Dim rec As DxVBLib.RECT 'Added for Direct3D -smchronos
Dim rec As DxVBLibA.RECT
Dim Vect_Position As D3DVECTOR2 'Added for Direct3D -smchronos

Vect_Position.X = MapItem(ItemNum).X * PIC_X
Vect_Position.Y = MapItem(ItemNum).Y * PIC_Y

    With rec
        .top = Item(MapItem(ItemNum).num).Pic * PIC_Y
        .bottom = .top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With
    Call Direct3D.DrawTex(itemtex, rec, Vect_Position)
End Sub

Public Sub BltFringeTile(ByVal X As Long, ByVal Y As Long)
Dim Fringe As Long
Dim Fringe2 As Long
Dim FringeAnim
'Dim rec As DxVBLib.RECT 'Added for Direct3D -smchronos
Dim rec As DxVBLibA.RECT
Dim Vect_Position As D3DVECTOR2 'Added for Direct3D -smchronos

Vect_Position.X = X * PIC_X
Vect_Position.Y = Y * PIC_Y
    
    Fringe = Map.Tile(X, Y).Fringe
    Fringe2 = Map.Tile(X, Y).Fringe2
    FringeAnim = Map.Tile(X, Y).FringeAnim
    
     If (MapAnim = 0) Or (FringeAnim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe > 0 Then
            With rec
                .top = Int(Fringe / 7) * PIC_Y
                .bottom = .top + PIC_Y
                .Left = (Fringe - Int(Fringe / 7) * 7) * PIC_X
                .Right = .Left + PIC_X
            End With
            Call Direct3D.DrawTex(tiletex, rec, Vect_Position)
            'Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If FringeAnim > 0 Then
            With rec
                .top = Int(FringeAnim / 7) * PIC_Y
                .bottom = .top + PIC_Y
                .Left = (FringeAnim - Int(FringeAnim / 7) * 7) * PIC_X
                .Right = .Left + PIC_X
            End With
            Call Direct3D.DrawTex(tiletex, rec, Vect_Position)
            'Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    
    If Fringe2 > 0 Then
        With rec
            .top = Int(Fringe2 / 7) * PIC_Y
            .bottom = .top + PIC_Y
            .Left = (Fringe2 - Int(Fringe2 / 7) * 7) * PIC_X
            .Right = .Left + PIC_X
        End With
        Call Direct3D.DrawTex(tiletex, rec, Vect_Position)
        'Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Public Sub BltPlayer(ByVal Index As Long, Optional ScaleSize As Single = 1, Optional Alpha As Long = &HFFFFFFFF)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
Dim Anim As Byte
Dim X As Long, Y As Long
'Dim rec As DxVBLib.RECT 'Added for Direct3D -smchronos
Dim rec As DxVBLibA.RECT
Dim Vect_Position As D3DVECTOR2 'Added for Direct3D -smchronos
    
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
    With Player(Index)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .top = GetPlayerSprite(Index) * PIC_Y
        .bottom = .top + PIC_Y
        .Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    X = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
    Y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
    
    Vect_Position.X = X
    Vect_Position.Y = Y
    
    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        With rec
            .top = .top + (Y * -1)
        End With
    End If
    
    If Player(Index).TintR = 0 And Player(Index).TintG = 0 And Player(Index).TintB = 0 Then
        Call Direct3D.DrawTex(spritetex, rec, Vect_Position, ScaleSize, Alpha)
    Else
        Call Direct3D.DrawTex(spritetex, rec, Vect_Position, ScaleSize, Direct3D.ConvertToHex(Player(Index).TintR, Player(Index).TintG, Player(Index).TintB))
    End If
End Sub

Sub BltPlayerDeath(ByVal Index As Long)
Dim Anim As Byte
Dim X As Long, Y As Long
'Dim rec As DxVBLib.RECT 'Added for Direct3D -smchronos
Dim rec As DxVBLibA.RECT
Dim Vect_Position As D3DVECTOR2 'Added for Direct3D -smchronos
    
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
    With Player(Index)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .top = GetPlayerSprite(Index) * PIC_Y
        .bottom = .top + PIC_Y
        .Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    X = Player(Index).Death.DeathX * PIC_X + Player(Index).XOffset
    Y = Player(Index).Death.DeathY * PIC_Y + Player(Index).YOffset - 4
    
    Vect_Position.X = X
    Vect_Position.Y = Y
    
    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        With rec
            .top = .top + (Y * -1)
        End With
    End If
    
    ' Add one to the animation count
    Player(Index).Death.AnimCount = Player(Index).Death.AnimCount + 1
        
    ' Draw the player fading out
    Call Direct3D.DrawTex(spritetex, rec, Vect_Position, 1, Direct3D.ConvertToHex(255 - (Player(Index).Death.AnimCount * 40), 255 - (Player(Index).Death.AnimCount * 40), 255 - (Player(Index).Death.AnimCount * 40), 255 - (Player(Index).Death.AnimCount * 40)))
        
    ' Check if it needs to be reset after 5 cycles
    If Player(Index).Death.AnimCount > 5 Then
        Player(Index).Death.Display = False
        Player(Index).Death.AnimCount = 0
    End If
End Sub

Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim Size As Long

Size = Direct3D.FontSize
    
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = C_Brown
            Case 1
                Color = C_DarkGrey
            Case 2
                Color = C_Cyan
            Case 3
                Color = C_Blue
            Case 4
                Color = C_Pink
        End Select
    Else
        Color = C_BrightRed
    End If
    
    ' Set the location
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - Int(PIC_Y / 2) - 6
    If TextX < 0 Then TextX = 0
    If TextX + PIC_X > frmDualSolace.picScreen.ScaleWidth Then TextX = TextX
    If TextY < 0 Then TextY = 0
    
    ' Refresh the font settings for names
    Call Direct3D.SetupFont("Verdana", 12, False, False)
    
    ' Stroke the text  - Rezeyu
    Call Direct3D.DrawText(Trim$(GetPlayerName(Index)), C_Black, DT_LEFT, TextX + 1, TextY)
    Call Direct3D.DrawText(Trim$(GetPlayerName(Index)), C_Black, DT_LEFT, TextX - 1, TextY)
    Call Direct3D.DrawText(Trim$(GetPlayerName(Index)), C_Black, DT_LEFT, TextX, TextY + 1)
    Call Direct3D.DrawText(Trim$(GetPlayerName(Index)), C_Black, DT_LEFT, TextX, TextY - 1)
    
    ' Draw the name
    Call Direct3D.DrawText(Trim$(GetPlayerName(Index)), Color, DT_LEFT, TextX, TextY)
End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Anim As Byte
Dim X As Long, Y As Long
'Dim rec As DxVBLib.RECT 'Added for Direct3D -smchronos
Dim rec As DxVBLibA.RECT
Dim Vect_Position As D3DVECTOR2 'Added for Direct3D -smchronos

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then
        Exit Sub
    End If
    
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
    With MapNpc(MapNpcNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
        .bottom = .top + PIC_Y
        .Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    With MapNpc(MapNpcNum)
        X = .X * PIC_X + .XOffset
        Y = .Y * PIC_Y + .YOffset - 4
    End With
    
    Vect_Position.X = X
    Vect_Position.Y = Y
    
    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        With rec
            .top = .top + (Y * -1)
        End With
    End If

    If Npc(MapNpc(MapNpcNum).num).TintR = 0 And Npc(MapNpc(MapNpcNum).num).TintG = 0 And Npc(MapNpc(MapNpcNum).num).TintB = 0 Then
        Call Direct3D.DrawTex(spritetex, rec, Vect_Position, 1, &HFFFFFFFF)
    Else
        Call Direct3D.DrawTex(spritetex, rec, Vect_Position, 1, Direct3D.ConvertToHex(Npc(MapNpc(MapNpcNum).num).TintR, Npc(MapNpc(MapNpcNum).num).TintG, Npc(MapNpc(MapNpcNum).num).TintB))
    End If
End Sub


Public Sub BltNpcDeath(ByVal MapNpcNum As Long)
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************

Dim Anim As Byte
Dim X As Long, Y As Long
'Dim rec As DxVBLib.RECT 'Added for Direct3D -smchronos
Dim rec As DxVBLibA.RECT
Dim Vect_Position As D3DVECTOR2 'Added for Direct3D -smchronos

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).num <= 0 Then
        Exit Sub
    End If
    
    ' Set animation
    Anim = 1
    
    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .top = Npc(MapNpc(MapNpcNum).num).Sprite * PIC_Y
        .bottom = .top + PIC_Y
        .Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    With MapNpc(MapNpcNum)
        X = .Death.DeathX * PIC_X + .XOffset
        Y = .Death.DeathY * PIC_Y + .YOffset - 4
    End With
    
    Vect_Position.X = X
    Vect_Position.Y = Y
    
    ' Check if its out of bounds because of the offset
    If Y < 0 Then
        Y = 0
        With rec
            .top = .top + (Y * -1)
        End With
    End If

    ' Add one to the animation count
    MapNpc(MapNpcNum).Death.AnimCount = MapNpc(MapNpcNum).Death.AnimCount + 1
        
    ' Draw the mapnpc fading out
    Call Direct3D.DrawTex(spritetex, rec, Vect_Position, 1, Direct3D.ConvertToHex(255 - (MapNpc(MapNpcNum).Death.AnimCount * 40), 255 - (MapNpc(MapNpcNum).Death.AnimCount * 40), 255 - (MapNpc(MapNpcNum).Death.AnimCount * 40), 255 - (MapNpc(MapNpcNum).Death.AnimCount * 40)))
        
    ' Check if it needs to be reset after 5 cycles
    If MapNpc(MapNpcNum).Death.AnimCount > 5 Then
        MapNpc(MapNpcNum).Death.Display = False
        MapNpc(MapNpcNum).Death.AnimCount = 0
    End If
End Sub

Sub BltNpcName(ByVal MapNpcNum As Long)
Dim TextX As Long
Dim TextY As Long

With Npc(MapNpc(MapNpcNum).num)
    ' Set the location
    TextX = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset + CLng(PIC_X / 2) - ((Len(Trim$(.Name)) / 2) * 8)
    TextY = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset - CLng(PIC_Y / 2) - 6
    
    ' Refresh the font settings for names
    Call Direct3D.SetupFont("Verdana", 12, False, False)
    
        ' Stroke the text  - Rezeyu
    Call Direct3D.DrawText(Trim$(.Name), C_Black, DT_LEFT, TextX + 1, TextY)
    Call Direct3D.DrawText(Trim$(.Name), C_Black, DT_LEFT, TextX - 1, TextY)
    Call Direct3D.DrawText(Trim$(.Name), C_Black, DT_LEFT, TextX, TextY + 1)
    Call Direct3D.DrawText(Trim$(.Name), C_Black, DT_LEFT, TextX, TextY - 1)
    
    ' Draw the name
    Call Direct3D.DrawText(Trim$(.Name), C_White, DT_LEFT, TextX, TextY)
End With
End Sub

Sub ProcessMovement(ByVal Index As Long)
    ' Check if player is walking, and if so process moving them over
    If Player(Index).Moving = MOVING_WALKING Then
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
            Player(Index).Moving = 0
        End If
    End If

    ' Check if player is running, and if so process moving them over
    If Player(Index).Moving = MOVING_RUNNING Then
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
            Player(Index).Moving = 0
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if player is walking, and if so process moving them over
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
        ' Broadcast message
        If Mid(frmDualSolace.txtChatEnter.Text, 1, 1) = "'" Then
            ChatText = Mid(frmDualSolace.txtChatEnter.Text, 2, Len(frmDualSolace.txtChatEnter.Text) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' Emote message
        If Mid(frmDualSolace.txtChatEnter.Text, 1, 1) = "-" Then
            ChatText = Mid(frmDualSolace.txtChatEnter.Text, 2, Len(frmDualSolace.txtChatEnter.Text) - 1)
            If Len(Trim$(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' Player message
        If Mid(frmDualSolace.txtChatEnter.Text, 1, 1) = "!" Then
            ChatText = Mid(frmDualSolace.txtChatEnter.Text, 2, Len(frmDualSolace.txtChatEnter.Text) - 1)
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
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
            
        ' // Commands //
        ' Help
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 5)) = "/help" Then
            Call AddText("Social Commands:", HelpColor)
            Call AddText("'msghere = Broadcast Message", HelpColor)
            Call AddText("-msghere = Emote Message", HelpColor)
            Call AddText("!namehere msghere = Player Message", HelpColor)
            Call AddText("Available Commands: /help, /info, /who, /fps, /inv, /stats, /train, /trade, /party, /join, /leave", HelpColor)
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' Verification User
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 5)) = "/info" Then
            ChatText = Mid(frmDualSolace.txtChatEnter.Text, 6, Len(frmDualSolace.txtChatEnter.Text) - 5)
            Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' Whos Online
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 4)) = "/who" Then
            Call SendWhosOnline
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
                        
        ' Checking fps
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 4)) = "/fps" Then
            Call AddText("FPS: " & GameFPS, Pink)
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
                
        ' Show inventory
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 4)) = "/inv" Then
            Call UpdateInventory
            frmDualSolace.picInv.Visible = True
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' Request stats
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
    
        ' Show training
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 6)) = "/train" Then
            If frmDualSolace.picTrainMenu.Visible = True Then
                frmDualSolace.picTrainMenu.Visible = False
            Else
                frmDualSolace.picTrainMenu.Visible = True
                frmDualSolace.cmbStat.Clear
                frmDualSolace.cmbStat.AddItem STRING_STRENGTH
                frmDualSolace.cmbStat.AddItem STRING_DEFENSE
                frmDualSolace.cmbStat.AddItem STRING_MAGIC
                frmDualSolace.cmbStat.AddItem STRING_SPEED
            End If
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If

        ' Request stats
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 6)) = "/trade" Then
            Call SendData("trade" & SEP_CHAR & END_CHAR)
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' Party request
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 6)) = "/party" Then
            ' Make sure they are actually sending something
            If Len(frmDualSolace.txtChatEnter.Text) > 7 Then
                ChatText = Mid(frmDualSolace.txtChatEnter.Text, 8, Len(frmDualSolace.txtChatEnter.Text) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Usage: /party playernamehere", AlertColor)
            End If
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' Join party
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 5)) = "/join" Then
            Call SendJoinParty
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' Leave party
        If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            frmDualSolace.txtChatEnter.Text = ""
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' Admin Help
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 6)) = "/admin" Then
                Call AddText("Social Commands:", HelpColor)
                Call AddText("""msghere = Global Admin Message", HelpColor)
                Call AddText("=msghere = Private Admin Message", HelpColor)
                Call AddText("Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /ban, /editspell", HelpColor)
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Kicking a player
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 5)) = "/kick" Then
                If Len(frmDualSolace.txtChatEnter.Text) > 6 Then
                    frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 7, Len(frmDualSolace.txtChatEnter.Text) - 6)
                    Call SendKick(frmDualSolace.txtChatEnter.Text)
                End If
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
        
            ' Global Message
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 1) = """" Then
                ChatText = Mid(frmDualSolace.txtChatEnter.Text, 2, Len(frmDualSolace.txtChatEnter.Text) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
        
            ' Admin Message
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 1) = "=" Then
                ChatText = Mid(frmDualSolace.txtChatEnter.Text, 2, Len(frmDualSolace.txtChatEnter.Text) - 1)
                If Len(Trim$(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
        End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
            ' Location
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Map Editor
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Warping to a player
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 9)) = "/warpmeto" Then
                If Len(frmDualSolace.txtChatEnter.Text) > 10 Then
                    frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 10, Len(frmDualSolace.txtChatEnter.Text) - 9)
                    Call WarpMeTo(frmDualSolace.txtChatEnter.Text)
                End If
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
                        
            ' Warping a player to you
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 9)) = "/warptome" Then
                If Len(frmDualSolace.txtChatEnter.Text) > 10 Then
                    frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 10, Len(frmDualSolace.txtChatEnter.Text) - 9)
                    Call WarpToMe(frmDualSolace.txtChatEnter.Text)
                End If
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
                        
            ' Warping to a map
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 7)) = "/warpto" Then
                If Len(frmDualSolace.txtChatEnter.Text) > 8 Then
                    frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 8, Len(frmDualSolace.txtChatEnter.Text) - 7)
                    n = Val(frmDualSolace.txtChatEnter.Text)
                
                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If
                End If
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Setting sprite
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 10)) = "/setsprite" Then
                If Len(frmDualSolace.txtChatEnter.Text) > 11 Then
                    ' Get sprite #
                    frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 12, Len(frmDualSolace.txtChatEnter.Text) - 11)
                
                    Call SendSetSprite(Val(frmDualSolace.txtChatEnter.Text))
                End If
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Map report
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
        
            ' Respawn request
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
        
            ' MOTD change
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 5) = "/motd" Then
                If Len(frmDualSolace.txtChatEnter.Text) > 6 Then
                    frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 7, Len(frmDualSolace.txtChatEnter.Text) - 6)
                    If Trim$(frmDualSolace.txtChatEnter.Text) <> "" Then
                        Call SendMOTDChange(frmDualSolace.txtChatEnter.Text)
                    End If
                End If
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Check the ban list
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 3) = "/banlist" Then
                Call SendBanList
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Banning a player
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 4)) = "/ban" Then
                If Len(frmDualSolace.txtChatEnter.Text) > 5 Then
                    frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 6, Len(frmDualSolace.txtChatEnter.Text) - 5)
                    Call SendBan(frmDualSolace.txtChatEnter.Text)
                    frmDualSolace.txtChatEnter.Text = ""
                End If
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Editing npc request
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 8) = "/editnpc" Then
                Call SendRequestEditNpc
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Editing shop request
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
        
            ' Editing spell request
            If Mid(frmDualSolace.txtChatEnter.Text, 1, 10) = "/editspell" Then
                Call SendRequestEditSpell
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = Val(Mid(frmDualSolace.txtChatEnter.Text, 12, 1))
                
                frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 14, Len(frmDualSolace.txtChatEnter.Text) - 13)
                
                Call SendSetAccess(frmDualSolace.txtChatEnter.Text, i)
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase(Mid(frmDualSolace.txtChatEnter.Text, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                frmDualSolace.txtChatEnter.Text = ""
                Exit Sub
            End If
        End If
        
        ' Say message
        If Len(Trim$(frmDualSolace.txtChatEnter.Text)) > 0 Then
            Call SayMsg(frmDualSolace.txtChatEnter.Text)
        End If
        frmDualSolace.txtChatEnter.Text = ""
        Exit Sub
    End If
    
    'Not needed!
    ' Handle when the user presses the backspace key
    'If (KeyAscii = vbKeyBack) Then
    '    If Len(frmDualSolace.txtChatEnter.Text) > 0 Then
    '        frmDualSolace.txtChatEnter.Text = Mid(frmDualSolace.txtChatEnter.Text, 1, Len(frmDualSolace.txtChatEnter.Text) - 1)
    '    End If
    'End If
    
    '* No longer using *
    ' And if neither, then add the character to the user's text buffer
    'If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) Then
        ' Make sure they just use standard keys, no macro keys
        'If KeyAscii >= 32 And KeyAscii <= 126 Then
            'frmDualSolace.txtChatEnter.Text = frmDualSolace.txtChatEnter.Text & Chr(KeyAscii)
        'End If
    'End If
    '* No longer using *
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim$(frmDualSolace.txtChatEnter.Text) = "" Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & SEP_CHAR & END_CHAR)
    End If
End Sub

Public Sub CheckAttack()
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Optimized function.
'****************************************************************
Dim X As Long, Y As Long
    
    If ControlDown = True And Player(MyIndex).AttackTimer + 1000 < GetTickCount And Player(MyIndex).Attacking = 0 Then
        With Player(MyIndex)
            .Attacking = 1
            .AttackTimer = GetTickCount
        End With
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
                If GetPlayerAccess(MyIndex) > 0 Then
                    Call SendData("ADMINPANEL" & SEP_CHAR & END_CHAR)
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
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Blocked = True Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).South = True Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).North = True Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Key = True Then
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
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Blocked = True Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).North = True Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).South = True Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Key = True Then
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
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Blocked = True Or Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).East = True Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).West = True Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Key = True Then
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
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Blocked = True Or Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).West = True Or Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).East = True Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Key = True Then
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
                If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Warp = True Then
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
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added gfx constants.
'****************************************************************
    
    SaveMap = Map
    InEditor = True
    frmDualSolace.picMapEditor.Visible = True
    Call SetPicSize(App.Path + GFX_PATH + "tiles" + GFX_EXT, frmDualSolace.picBackSelect)
    'With frmDualSolace.picBack
    '    .Width = PicSize.Width
    '    .Height = PicSize.Height
    'End With
    With frmDualSolace.picBackSelect
        .Picture = LoadPicture(App.Path + GFX_PATH + "tiles" + GFX_EXT)
        frmDualSolace.scrlPicture.Max = Int(.Height \ PIC_Y) - 1
    End With
    'Debug.Print frmDualSolace.scrlPicture.Max
    
    'Load map attribute graphic as texture
    
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1, y1, n1, n2, n3, n4 As Long
Dim x2 As Long, y2 As Long

    If InEditor Then
        x1 = Int(X / PIC_X)
        y1 = Int(Y / PIC_Y)
        'x1 = CLng(frmDualSolace.lblPosX.Caption)
        'y1 = CLng(frmDualSolace.lblPosY.Caption)
        If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmDualSolace.shpSelected.Height <= 32 And frmDualSolace.shpSelected.Width <= 32 Then
            If frmDualSolace.optLayers.Value = True Then
                    With Map.Tile(x1, y1)
                        If frmDualSolace.optGround.Value = True Then .Ground = EditorTileY * 7 + EditorTileX
                        If frmDualSolace.optMask.Value = True Then .Mask = EditorTileY * 7 + EditorTileX
                        If frmDualSolace.optMask2.Value = True Then .Mask2 = EditorTileY * 7 + EditorTileX
                        If frmDualSolace.optAnim.Value = True Then .Anim = EditorTileY * 7 + EditorTileX
                        If frmDualSolace.optAnim2.Value = True Then .Anim2 = EditorTileY * 7 + EditorTileX
                        If frmDualSolace.optFringe.Value = True Then .Fringe = EditorTileY * 7 + EditorTileX
                        If frmDualSolace.optFringeAnim.Value = True Then .FringeAnim = EditorTileY * 7 + EditorTileX
                        If frmDualSolace.optFringe2.Value = True Then .Fringe2 = EditorTileY * 7 + EditorTileX
                    End With
            Else
                With Map.Tile(x1, y1)
                    If frmDualSolace.chkBlocked.Value = Checked Then
                        .Blocked = True
                        .Walkable = False
                    Else
                        .Blocked = False
                        .Walkable = True
                    End If
                    If frmDualSolace.chkNorth.Value = Checked Then
                        .North = True
                    Else
                        .North = False
                    End If
                    If frmDualSolace.chkWest.Value = Checked Then
                        .West = True
                    Else
                        .West = False
                    End If
                    If frmDualSolace.chkEast.Value = Checked Then
                        .East = True
                    Else
                        .East = False
                    End If
                    If frmDualSolace.chkSouth.Value = Checked Then
                        .South = True
                    Else
                        .South = False
                    End If
                    If frmDualSolace.chkWarp.Value = Checked Then
                        .Warp = True
                        .WarpMap = EditorWarpMap
                        .WarpX = EditorWarpX
                        .WarpY = EditorWarpY
                    Else
                        .Warp = False
                        .WarpMap = 0
                        .WarpX = 0
                        .WarpY = 0
                    End If
                    If frmDualSolace.chkItem.Value = Checked Then
                        .Item = True
                        .ItemNum = ItemEditorNum
                        .ItemValue = ItemEditorValue
                    Else
                        .Item = False
                        .ItemNum = 0
                        .ItemValue = 0
                    End If
                    If frmDualSolace.chkNpcAvoid.Value = Checked Then
                        .NpcAvoid = True
                    Else
                        .NpcAvoid = False
                    End If
                    If frmDualSolace.chkKey.Value = Checked Then
                        .Key = True
                        .KeyNum = KeyEditorNum
                        .KeyTake = KeyEditorTake
                    Else
                        .Key = False
                        .KeyNum = 0
                        .KeyTake = 0
                    End If
                    If frmDualSolace.chkKeyOpen.Value = Checked Then
                        .KeyOpen = True
                        .KeyOpenX = KeyOpenEditorX
                        .KeyOpenY = KeyOpenEditorY
                    Else
                        .KeyOpen = False
                        .KeyOpenX = 0
                        .KeyOpenY = 0
                    End If
                    If frmDualSolace.chkBank.Value = Checked Then
                        .Bank = True
                    Else
                        .Bank = False
                    End If
                    If frmDualSolace.chkShop.Value = Checked Then
                        .Shop = True
                        .ShopNum = EditorShopNum
                    Else
                        .Shop = False
                        .ShopNum = 0
                    End If
                    If frmDualSolace.chkHeal.Value = Checked Then
                        .Heal = True
                        .HealValue = EditorHealValue
                    Else
                        .Heal = False
                        .HealValue = 0
                    End If
                    If frmDualSolace.chkDamage.Value = Checked Then
                        .Damage = True
                        .DamageValue = EditorDamageValue
                    Else
                        .Damage = False
                        .DamageValue = 0
                    End If
                End With
            End If
            Else
                  For y2 = 0 To Int(frmDualSolace.shpSelected.Height / PIC_Y) - 1
                      For x2 = 0 To Int(frmDualSolace.shpSelected.Width / PIC_X) - 1
                            If x1 + x2 <= MAX_MAPX Then
                                If y1 + y2 <= MAX_MAPY Then
                                    If frmDualSolace.optLayers.Value = True Then
                                          With Map.Tile(x1 + x2, y1 + y2)
                                              If frmDualSolace.optGround.Value = True Then .Ground = (EditorTileY + y2) * 256 + (EditorTileX + x2)
                                              If frmDualSolace.optMask.Value = True Then .Mask = (EditorTileY + y2) * 256 + (EditorTileX + x2)
                                              If frmDualSolace.optMask2.Value = True Then .Mask2 = (EditorTileY + y2) * 256 + (EditorTileX + x2)
                                              If frmDualSolace.optAnim.Value = True Then .Anim = (EditorTileY + y2) * 256 + (EditorTileX + x2)
                                              If frmDualSolace.optAnim2.Value = True Then .Anim2 = (EditorTileY + y2) * 256 + (EditorTileX + x2)
                                              If frmDualSolace.optFringe.Value = True Then .Fringe = (EditorTileY + y2) * 256 + (EditorTileX + x2)
                                              If frmDualSolace.optFringeAnim.Value = True Then .FringeAnim = (EditorTileY + y2) * 256 + (EditorTileX + x2)
                                              If frmDualSolace.optFringe2.Value = True Then .Fringe2 = (EditorTileY + y2) * 256 + (EditorTileX + x2)
                                          End With
                                    End If
                                End If
                            End If
                      Next x2
                  Next y2
              End If
        End If
        
        If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmDualSolace.optLayers.Value = True Then
                With Map.Tile(x1, y1)
                    If frmDualSolace.optGround.Value = True Then .Ground = 0
                    If frmDualSolace.optMask.Value = True Then .Mask = 0
                    If frmDualSolace.optMask2.Value = True Then .Mask2 = 0
                    If frmDualSolace.optAnim.Value = True Then .Anim = 0
                    If frmDualSolace.optAnim2.Value = True Then .Anim2 = 0
                    If frmDualSolace.optFringe.Value = True Then .Fringe = 0
                    If frmDualSolace.optFringeAnim.Value = True Then .FringeAnim = 0
                    If frmDualSolace.optFringe2.Value = True Then .Fringe2 = 0
                End With
            Else
                With Map.Tile(x1, y1)
                    .Walkable = True
                    .Blocked = False
                    .North = False
                    .West = False
                    .East = False
                    .South = False
                    .Warp = False
                    .WarpMap = 0
                    .WarpX = 0
                    .WarpY = 0
                    .Item = False
                    .ItemNum = 0
                    .ItemValue = 0
                    .NpcAvoid = False
                    .Key = False
                    .KeyNum = 0
                    .KeyTake = 0
                    .KeyOpen = False
                    .KeyOpenX = 0
                    .KeyOpenY = 0
                    .Bank = False
                    .Shop = False
                    .ShopNum = 0
                    .Heal = False
                    .HealValue = 0
                    .Damage = False
                    .DamageValue = 0
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
    frmDualSolace.shpSelected.top = Int(EditorTileY * PIC_Y)
    frmDualSolace.shpSelected.Left = Int(EditorTileX * PIC_Y)
    Call BitBlt(frmDualSolace.picSelect.hDC, 0, 0, PIC_X, PIC_Y, frmDualSolace.picBackSelect.hDC, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub EditorChooseTiles(X As Single, Y As Single)
    If EditorTileX <> Int(X \ PIC_X) Or EditorTileY <> Int(Y \ PIC_Y) Then
        If EditorTileX < Int(X \ PIC_X) Or EditorTileY < Int(Y \ PIC_Y) Then
            EditorTileXEnd = Int(X \ PIC_X)
            EditorTileYEnd = Int(Y \ PIC_Y)
            
        ElseIf EditorTileX > Int(X \ PIC_X) Or EditorTileY > Int(Y \ PIC_Y) Then
            EditorTileX = Int(X \ PIC_X)
            EditorTileY = Int(Y \ PIC_Y)
            EditorTileXEnd = EditorTileX
            EditorTileYEnd = EditorTileY
        End If
    End If
End Sub

Public Sub EditorTileScroll()
    frmDualSolace.picBackSelect.top = (frmDualSolace.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    Map = SaveMap
    InEditor = False
    frmDualSolace.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, X As Long, Y As Long

    ' Ground layer
    If frmDualSolace.optGround.Value = True Then
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
    If frmDualSolace.optMask.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Mask = 0
                Next X
            Next Y
        End If
    End If
    
    ' Mask2 layer
    If frmDualSolace.optMask2.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the second mask layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Mask2 = 0
                Next X
            Next Y
        End If
    End If

    ' Animation layer
    If frmDualSolace.optAnim.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Anim = 0
                Next X
            Next Y
        End If
    End If
    
    ' Animation2 layer
    If frmDualSolace.optAnim2.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the second animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Anim2 = 0
                Next X
            Next Y
        End If
    End If

    ' Fringe layer
    If frmDualSolace.optFringe.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Fringe = 0
                Next X
            Next Y
        End If
    End If
    
    ' FringeAnim layer
    If frmDualSolace.optFringeAnim.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).FringeAnim = 0
                Next X
            Next Y
        End If
    End If
    
    ' Fringe2 layer
    If frmDualSolace.optFringe2.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the second fringe layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For Y = 0 To MAX_MAPY
                For X = 0 To MAX_MAPX
                    Map.Tile(X, Y).Fringe2 = 0
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
                With Map.Tile(X, Y)
                    .Walkable = True
                    .Blocked = False
                    .North = False
                    .West = False
                    .East = False
                    .South = False
                    .Warp = False
                    .WarpMap = 0
                    .WarpX = 0
                    .WarpY = 0
                    .Item = False
                    .ItemNum = 0
                    .ItemValue = 0
                    .NpcAvoid = False
                    .Key = False
                    .KeyNum = 0
                    .KeyTake = 0
                    .KeyOpen = False
                    .KeyOpenX = 0
                    .KeyOpenY = 0
                    .Bank = False
                    .Shop = False
                    .ShopNum = 0
                    .Heal = False
                    .HealValue = 0
                    .Damage = False
                    .DamageValue = 0
                End With
            Next X
        Next Y
    End If
End Sub

Public Sub ItemEditorInit()
Dim FSys As Object, Folder As Object, FolderFiles As Object, File As Object
Dim FileHold As String, FileHold2 As String
Dim X As Long
Set FSys = CreateObject("Scripting.FileSystemObject")

    Call SetPicSize(App.Path + GFX_PATH + "items" + GFX_EXT, frmEditor.picItems)

    frmEditor.picItems.Picture = LoadPicture(App.Path & GFX_PATH & "items" & GFX_EXT)
    
    frmEditor.txtItemName.Text = Trim$(Item(EditorIndex).Name)
    frmEditor.txtItemDescription.Text = Trim$(Item(EditorIndex).Description)
    frmEditor.scrlItemPic.Value = Item(EditorIndex).Pic
    frmEditor.cmbItemType.ListIndex = Item(EditorIndex).Type
    
    If Item(EditorIndex).Data4 = WEAPON_TYPE_BOW Then
        frmEditor.chkBow.Value = Checked
    Else
        frmEditor.chkBow.Value = Unchecked
    End If
    
    If Item(EditorIndex).Data4 = SHIELD_TYPE_ARROW Then
        frmEditor.chkArrow.Value = Checked
    Else
        frmEditor.chkArrow.Value = Unchecked
    End If
    
    If Item(EditorIndex).Data5 = 1 Then
        frmEditor.chkUnBreakable.Value = Checked
    Else
        frmEditor.chkUnBreakable.Value = Unchecked
    End If
    
    If (frmEditor.cmbItemType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor.cmbItemType.ListIndex <= ITEM_TYPE_SHIELD) Then
        frmEditor.fraItemEquipment.Visible = True
        frmEditor.scrlItemDurability.Value = Item(EditorIndex).Data1
        frmEditor.scrlItemStrength.Value = Item(EditorIndex).Data2
    Else
        frmEditor.fraItemEquipment.Visible = False
    End If
    
    If (frmEditor.cmbItemType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmEditor.cmbItemType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmEditor.fraItemVitals.Visible = True
        frmEditor.scrlItemVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmEditor.fraItemVitals.Visible = False
    End If
    
    If (frmEditor.cmbItemType.ListIndex = ITEM_TYPE_SPELL) Then
        frmEditor.fraItemSpell.Visible = True
        frmEditor.scrlItemSpell.Value = Item(EditorIndex).Data1
    Else
        frmEditor.fraItemSpell.Visible = False
    End If
    
    frmEditor.scrlItemPic.Max = (Int(Direct3D.GetPicHeight(App.Path + GFX_PATH + "items" + GFX_EXT) \ 32) - 1)

    'Load sound list, find sound set
    'Set the folder objects
    Set Folder = FSys.GetFolder(App.Path & "\sound")
    Set FolderFiles = Folder.Files
    
    frmEditor.cmbSound.Clear
    frmEditor.cmbSound.AddItem "No Sound"
    frmEditor.cmbSound.ListIndex = 0
    For Each File In FolderFiles
        FileHold = Mid(File, Len(App.Path & "\sound\") + 1, ((Len(File) - Len(App.Path & "\sound\"))))
        FileHold2 = Mid(FileHold, Len(FileHold) - Len(".wav") + 1)
        Debug.Print "Sound file: " & FileHold
        Debug.Print "Sound name: " & LCase$(FileHold2)
        If LCase$(FileHold2) = ".wav" Then
            frmEditor.cmbSound.AddItem Mid(FileHold, 1, Len(FileHold) - Len(".wav"))
        End If
    Next File
    
    For X = 0 To frmEditor.cmbSound.ListCount - 1
        If LCase$(frmEditor.cmbSound.List(X)) = LCase$(Item(EditorIndex).Sound) Then
            frmEditor.cmbSound.ListIndex = X
            Exit Sub
        End If
    Next X
    
    If frmEditor.cmbSound.ListIndex = -1 Then frmEditor.cmbSound.ListIndex = 0
    
    'Destroy the folder objects
    Set File = Nothing
    Set FolderFiles = Nothing
    Set Folder = Nothing
    Set FSys = Nothing
End Sub

Public Sub ItemEditorSave()
    Item(EditorIndex).Name = frmEditor.txtItemName.Text
    Item(EditorIndex).Description = frmEditor.txtItemDescription.Text
    Item(EditorIndex).Pic = frmEditor.scrlItemPic.Value
    Item(EditorIndex).Type = frmEditor.cmbItemType.ListIndex

    If (frmEditor.cmbItemType.ListIndex >= ITEM_TYPE_WEAPON) And (frmEditor.cmbItemType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmEditor.scrlItemDurability.Value
        Item(EditorIndex).Data2 = frmEditor.scrlItemStrength.Value
        Item(EditorIndex).Data3 = 0
        
        If frmEditor.chkBow.Value = Checked And frmEditor.cmbItemType.ListIndex = ITEM_TYPE_WEAPON Then
            Item(EditorIndex).Data4 = WEAPON_TYPE_BOW
        ElseIf frmEditor.chkArrow.Value = Checked And frmEditor.cmbItemType.ListIndex = ITEM_TYPE_SHIELD Then
            Item(EditorIndex).Data4 = SHIELD_TYPE_ARROW
        Else
            Item(EditorIndex).Data4 = 0
        End If
        
        If frmEditor.chkUnBreakable.Value = Checked Then
            Item(EditorIndex).Data5 = 1
        Else
            Item(EditorIndex).Data5 = 0
        End If
    End If
    
    If (frmEditor.cmbItemType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmEditor.cmbItemType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmEditor.scrlItemVitalMod.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).Data4 = 0
        Item(EditorIndex).Data5 = 0
    End If
    
    If (frmEditor.cmbItemType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmEditor.scrlItemSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).Data4 = 0
        Item(EditorIndex).Data5 = 0
    End If
    
    'Set sound
    Item(EditorIndex).Sound = Trim$(frmEditor.cmbSound.List(frmEditor.cmbSound.ListIndex))
    
    Call SendSaveItem(EditorIndex)
End Sub

Public Sub ItemEditorBltItem()
    Call BitBlt(frmEditor.picItemPic.hDC, 0, 0, PIC_X, PIC_Y, frmEditor.picItems.hDC, 0, frmEditor.scrlItemPic.Value * PIC_Y, SRCCOPY)
End Sub

Sub DrawInventory()
'On Error GoTo Pic
'Dim sRect As DxVBLib.RECT
'Dim dRect As DxVBLib.RECT
Dim sRect As RECT
Dim dRect As RECT
Dim n As Byte, lCount As Long, Slot As Byte, Value As String, TotalK As Long
        
        dRect.top = 0
        dRect.Left = 0
        dRect.Right = dRect.Left + PIC_X
        dRect.bottom = dRect.top + PIC_Y
        TotalK = 0
        lCount = 0
        
    For n = 0 To (MAX_INV - 1)
        If Player(MyIndex).Inv(n + 1).num > 0 Then
            sRect.top = (Item(Player(MyIndex).Inv(n + 1).num).Pic * PIC_Y)
            sRect.Left = 0
            sRect.Right = sRect.Left + PIC_X
            sRect.bottom = sRect.top + PIC_Y
            Call BitBlt(frmDualSolace.picItem(n).hDC, dRect.Left, dRect.top, PIC_X, PIC_Y, frmEditor.picItems.hDC, sRect.Left, sRect.top, SRCCOPY)
            frmDualSolace.picItem(n).Refresh
            DoEvents
            'Call DD_ItemSurf.BltToDC(frmDualSolace.picItem(n).hdc, sRect, dRect)

            If GetPlayerInvItemValue(MyIndex, n + 1) >= 1000 Then
                For lCount = 1 To Int((GetPlayerInvItemValue(MyIndex, n + 1) \ 1000))
                    TotalK = TotalK + 1
                Next lCount
                Value = CStr(TotalK) & "K+"
            Else
                Value = CStr(GetPlayerInvItemValue(MyIndex, n + 1))
            End If
            
            If Item(Player(MyIndex).Inv(n + 1).num).Type = ITEM_TYPE_CURRENCY Then
                frmDualSolace.picItem(n).ToolTipText = "Name: " & Trim$(Item(GetPlayerInvItemNum(MyIndex, n + 1)).Name) & " | " & _
                                                       "Total: " & CStr(GetPlayerInvItemValue(MyIndex, n + 1))
            Else
                frmDualSolace.picItem(n).ToolTipText = "Name: " & Trim$(Item(GetPlayerInvItemNum(MyIndex, n + 1)).Name) & " | " & _
                                                       "Durability: " & CStr(GetPlayerInvItemDur(MyIndex, n + 1))
            End If
            If frmDualSolace.picBank.Visible = True Then
                Call BitBlt(frmDualSolace.picBankInv(n).hDC, dRect.Left, dRect.top, PIC_X, PIC_Y, frmEditor.picItems.hDC, sRect.Left, sRect.top, SRCCOPY)
                frmDualSolace.picBankInv(n).Refresh
                DoEvents
                'Call DD_ItemSurf.BltToDC(frmDualSolace.picBankInv(n).hdc, sRect, dRect)
                'Call DrawText(frmDualSolace.picBankInv(n).hdc, 0, 16, Value, QBColor(White))
                'frmDualSolace.picBankInv(n).Refresh
            End If
        Else
            frmDualSolace.picItem(n).Picture = Nothing
            If frmDualSolace.picBank.Visible = True Then
                frmDualSolace.picBankInv(n).Picture = Nothing
                'frmDualSolace.picBankInv(n).Refresh
            End If
        End If
    Next n
    
    If frmDualSolace.picBank.Visible = True Then
        For n = 0 To 19
            Slot = (frmDualSolace.scrlBank.Value * 5)
            Slot = Slot + n
            
            Debug.Print "Slot: " & Slot
            
            'If n > 19 And n < 39 Then Slot = n - 19
            If Player(MyIndex).BankInv(Slot + 1).num > 0 Then
                sRect.top = (Item(Player(MyIndex).BankInv(Slot + 1).num).Pic * PIC_Y)
                sRect.Left = 0
                sRect.Right = sRect.Left + PIC_X
                sRect.bottom = sRect.top + PIC_Y
                Call BitBlt(frmDualSolace.picBankItem(n).hDC, dRect.Left, dRect.top, PIC_X, PIC_Y, frmEditor.picItems.hDC, sRect.Left, sRect.top, SRCCOPY)
                frmDualSolace.picBankItem(n).Refresh
                DoEvents
                'Call DD_ItemSurf.BltToDC(frmDualSolace.picBankItem(n).hdc, sRect, dRect)
                'Call DrawText(frmDualSolace.picBankItem(n).hdc, 0, 16, Value, QBColor(White))
                'frmDualSolace.picItem(n).Refresh
            Else
                frmDualSolace.picBankItem(n).Picture = Nothing
                frmDualSolace.picBankItem(n).Refresh
            End If
        Next n
    End If
    
    Exit Sub
Pic:
Exit Sub
End Sub

Sub DrawEquipment(ByVal Weapon As Long, ByVal WeaponDur As Long, ByVal Armor As Long, ByVal ArmorDur As Long, ByVal Helmet As Long, ByVal HelmetDur As Long, ByVal Shield As Long, ByVal ShieldDur As Long)
'On Error GoTo Pic
'Dim sRect As DxVBLib.RECT
'Dim dRect As DxVBLib.RECT
Dim sRect As RECT
Dim dRect As RECT
        
        dRect.top = 0
        dRect.Left = 0
        dRect.Right = dRect.Left + PIC_X
        dRect.bottom = dRect.top + PIC_Y
        
        If Weapon > 0 Then
            sRect.top = (Item(Weapon).Pic * PIC_Y)
            sRect.Left = 0
            sRect.Right = sRect.Left + PIC_X - 1
            sRect.bottom = sRect.top + PIC_Y - 1
            Call BitBlt(frmDualSolace.picWeapon.hDC, dRect.Left, dRect.top, PIC_X, PIC_Y, frmEditor.picItems.hDC, sRect.Left, sRect.top, SRCCOPY)
            frmDualSolace.picWeapon.Refresh
            DoEvents
            'Call DD_ItemSurf.BltToDC(frmDualSolace.picWeapon.hDC, sRect, dRect)
            frmDualSolace.lblWeapon.Caption = CStr(WeaponDur)
        Else
            frmDualSolace.picWeapon.Picture = Nothing
            frmDualSolace.lblWeapon.Caption = ""
        End If
        
        dRect.top = 0
        dRect.Left = 0
        dRect.Right = dRect.Left + PIC_X
        dRect.bottom = dRect.top + PIC_Y
        
        If Armor > 0 Then
            sRect.top = Item(Armor).Pic * PIC_Y
            sRect.Left = 0
            sRect.Right = sRect.Left + PIC_X
            sRect.bottom = sRect.top + PIC_Y
            Call BitBlt(frmDualSolace.picArmor.hDC, dRect.Left, dRect.top, PIC_X, PIC_Y, frmEditor.picItems.hDC, sRect.Left, sRect.top, SRCCOPY)
            frmDualSolace.picArmor.Refresh
            DoEvents
            'Call DD_ItemSurf.BltToDC(frmDualSolace.picArmor.hDC, sRect, dRect)
            frmDualSolace.lblArmor.Caption = CStr(ArmorDur)
        Else
            frmDualSolace.picArmor.Picture = Nothing
            frmDualSolace.lblArmor.Caption = ""
        End If
        
        dRect.top = 0
        dRect.Left = 0
        dRect.Right = dRect.Left + PIC_X
        dRect.bottom = dRect.top + PIC_Y
        
        If Helmet > 0 Then
            sRect.top = Item(Helmet).Pic * PIC_Y
            sRect.Left = 0
            sRect.Right = sRect.Left + PIC_X
            sRect.bottom = sRect.top + PIC_Y
            Call BitBlt(frmDualSolace.picHelmet.hDC, dRect.Left, dRect.top, PIC_X, PIC_Y, frmEditor.picItems.hDC, sRect.Left, sRect.top, SRCCOPY)
            frmDualSolace.picHelmet.Refresh
            DoEvents
            'Call DD_ItemSurf.BltToDC(frmDualSolace.picHelmet.hDC, sRect, dRect)
            frmDualSolace.lblHelmet.Caption = CStr(HelmetDur)
        Else
            frmDualSolace.picHelmet.Picture = Nothing
            frmDualSolace.lblHelmet.Caption = ""
        End If
        
        dRect.top = 0
        dRect.Left = 0
        dRect.Right = dRect.Left + PIC_X
        dRect.bottom = dRect.top + PIC_Y
        
        If Shield > 0 Then
            sRect.top = Item(Shield).Pic * PIC_Y
            sRect.Left = 0
            sRect.Right = sRect.Left + PIC_X
            sRect.bottom = sRect.top + PIC_Y
            Call BitBlt(frmDualSolace.picShield.hDC, dRect.Left, dRect.top, PIC_X, PIC_Y, frmEditor.picItems.hDC, sRect.Left, sRect.top, SRCCOPY)
            frmDualSolace.picShield.Refresh
            DoEvents
            'Call DD_ItemSurf.BltToDC(frmDualSolace.picShield.hDC, sRect, dRect)
            frmDualSolace.lblShield.Caption = CStr(ShieldDur)
        Else
            frmDualSolace.picShield.Picture = Nothing
            frmDualSolace.lblShield.Caption = ""
        End If
        
Pic:
Exit Sub
End Sub

Public Sub NpcEditorInit()
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/12/2005  Shannara   Added gfx constant.
'****************************************************************
    
    Call SetPicSize(App.Path + GFX_PATH + "sprites" + GFX_EXT, frmEditor.picSprites)
    
    frmEditor.picSprites.Picture = LoadPicture(App.Path & GFX_PATH & "sprites" & GFX_EXT)
    
    frmEditor.txtNpcName.Text = Trim$(Npc(EditorIndex).Name)
    frmEditor.txtAttackSay.Text = Trim$(Npc(EditorIndex).AttackSay)
    frmEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmEditor.txtSpawnSecs.Text = Trim$(CStr(Npc(EditorIndex).SpawnSecs))
    frmEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmEditor.txtChance.Text = Trim$(CStr(Npc(EditorIndex).DropChance))
    frmEditor.scrlNum.Value = Npc(EditorIndex).DropItem
    frmEditor.scrlValue.Value = Npc(EditorIndex).DropItemValue
    frmEditor.txtHP.Text = Trim$(CStr(Npc(EditorIndex).HP))
    frmEditor.txtEXP.Text = Trim$(CStr(Npc(EditorIndex).EXP))
    Debug.Print Npc(EditorIndex).HP
    Debug.Print Npc(EditorIndex).EXP
    frmEditor.scrlSTR.Value = Npc(EditorIndex).STR
    frmEditor.scrlDEF.Value = Npc(EditorIndex).DEF
    frmEditor.scrlSPEED.Value = Npc(EditorIndex).Speed
    frmEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmEditor.scrlR.Value = Npc(EditorIndex).TintR
    frmEditor.scrlG.Value = Npc(EditorIndex).TintG
    frmEditor.scrlB.Value = Npc(EditorIndex).TintB
    
    frmEditor.scrlSprite.Max = (Int(Direct3D.GetPicHeight(App.Path + GFX_PATH + "sprites" + GFX_EXT) \ 32) - 1)
    
    If Npc(EditorIndex).Fear = True Then
        frmEditor.chkAfraid.Value = Checked
    Else
        frmEditor.chkAfraid.Value = Unchecked
    End If
End Sub

Public Sub NpcEditorSave()
    Npc(EditorIndex).Name = frmEditor.txtNpcName.Text
    Npc(EditorIndex).AttackSay = frmEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmEditor.scrlSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmEditor.scrlRange.Value
    Npc(EditorIndex).DropChance = Val(frmEditor.txtChance.Text)
    Npc(EditorIndex).DropItem = frmEditor.scrlNum.Value
    Npc(EditorIndex).DropItemValue = frmEditor.scrlValue.Value
    Npc(EditorIndex).HP = CLng(frmEditor.txtHP.Text)
    Npc(EditorIndex).STR = frmEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmEditor.scrlDEF.Value
    Npc(EditorIndex).Speed = frmEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmEditor.scrlMAGI.Value
    Npc(EditorIndex).EXP = CLng(frmEditor.txtEXP.Text)
    
    If frmEditor.chkAfraid.Value = 1 Then
        Npc(EditorIndex).Fear = True
    Else
        Npc(EditorIndex).Fear = False
    End If
    
    ' Tint variables, need to find a good way to save them as a hex?
    Npc(EditorIndex).TintR = frmEditor.scrlR.Value
    Npc(EditorIndex).TintG = frmEditor.scrlG.Value
    Npc(EditorIndex).TintB = frmEditor.scrlB.Value
    
    Call SendSaveNpc(EditorIndex)
End Sub

Public Sub NpcEditorBltSprite()
    Call BitBlt(frmEditor.picSprite.hDC, 0, 0, PIC_X, PIC_Y, frmEditor.picSprites.hDC, 3 * PIC_X, frmEditor.scrlSprite.Value * PIC_Y, SRCCOPY)
End Sub

Public Sub ClassEditorBltSprite()
    Call BitBlt(frmEditor.picClassSprite.hDC, 0, 0, PIC_X, PIC_Y, frmEditor.picSprites.hDC, 3 * PIC_X, frmEditor.scrlClassSprite.Value * PIC_Y, SRCCOPY)
End Sub

Public Sub ShopEditorInit()
On Error Resume Next

Dim i As Long

    frmEditor.txtShopName.Text = Trim$(Shop(EditorIndex).Name)
    frmEditor.txtJoinSay.Text = Trim$(Shop(EditorIndex).JoinSay)
    frmEditor.txtLeaveSay.Text = Trim$(Shop(EditorIndex).LeaveSay)
    frmEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    
    If Shop(EditorIndex).Restock = TIME_MINUTE Then
        frmEditor.cmdRestock.ListIndex = 0
    ElseIf Shop(EditorIndex).Restock = TIME_HOUR Then
        frmEditor.cmdRestock.ListIndex = 1
    ElseIf Shop(EditorIndex).Restock = TIME_FULL Then
        frmEditor.cmdRestock.ListIndex = 2
    Else
        frmEditor.cmdRestock.ListIndex = 0
    End If
    
    frmEditor.cmbItemGive.Clear
    frmEditor.cmbItemGive.AddItem "None"
    frmEditor.cmbItemGet.Clear
    frmEditor.cmbItemGet.AddItem "None"
    For i = 1 To MAX_ITEMS
        frmEditor.cmbItemGive.AddItem i & ": " & Trim$(Item(i).Name)
        frmEditor.cmbItemGet.AddItem i & ": " & Trim$(Item(i).Name)
    Next i
    frmEditor.cmbItemGive.ListIndex = 0
    frmEditor.cmbItemGet.ListIndex = 0
    
    Call UpdateShopTrade
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long, Stock As Long
    
    frmEditor.lstTradeItem.Clear
    For i = 1 To MAX_TRADES
        GetItem = Shop(EditorIndex).TradeItem(i).GetItem
        GetValue = Shop(EditorIndex).TradeItem(i).GetValue
        GiveItem = Shop(EditorIndex).TradeItem(i).GiveItem
        GiveValue = Shop(EditorIndex).TradeItem(i).GiveValue
        Stock = Shop(EditorIndex).TradeItem(i).MaxStock
        
        If GetItem > 0 And GiveItem > 0 And Stock = -1 Then
            frmEditor.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name) & "(Infinite)"
        ElseIf GetItem > 0 And GiveItem > 0 And Stock > 0 Then
            frmEditor.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim$(Item(GiveItem).Name) & " for " & GetValue & " " & Trim$(Item(GetItem).Name) & "(" & Stock & ")"
        Else
            frmEditor.lstTradeItem.AddItem "Empty Trade Slot"
        End If
    Next i
    frmEditor.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorSave()
    Shop(EditorIndex).Name = frmEditor.txtShopName.Text
    Shop(EditorIndex).JoinSay = frmEditor.txtJoinSay.Text
    Shop(EditorIndex).LeaveSay = frmEditor.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmEditor.chkFixesItems.Value
    
    If frmEditor.cmdRestock.ListIndex = 0 Then
    Shop(EditorIndex).Restock = TIME_MINUTE
    ElseIf frmEditor.cmdRestock.ListIndex = 1 Then
    Shop(EditorIndex).Restock = TIME_HOUR
    ElseIf frmEditor.cmdRestock.ListIndex = 2 Then
    Shop(EditorIndex).Restock = TIME_FULL
    End If
    
    Call SendSaveShop(EditorIndex)
End Sub

Public Sub SpellEditorInit()
On Error Resume Next

Dim i As Long

    frmEditor.cmbSpellClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmEditor.cmbSpellClassReq.AddItem Trim$(Class(i).Name)
    Next i
    
    frmEditor.txtSpellName.Text = Trim$(Spell(EditorIndex).Name)
    frmEditor.cmbSpellClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmEditor.scrlSpellLevelReq.Value = Spell(EditorIndex).LevelReq
        
    frmEditor.cmbSpellType.ListIndex = Spell(EditorIndex).Type
    If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
        frmEditor.fraSpellVitals.Visible = True
        frmEditor.fraGiveItem.Visible = False
        frmEditor.scrlSpellVitalMod.Value = Spell(EditorIndex).Data1
    Else
        frmEditor.fraSpellVitals.Visible = False
        frmEditor.fraGiveItem.Visible = True
        frmEditor.scrlSpellItemNum.Value = Spell(EditorIndex).Data1
        frmEditor.scrlSpellItemValue.Value = Spell(EditorIndex).Data2
    End If
End Sub

Public Sub SpellEditorSave()
    Spell(EditorIndex).Name = frmEditor.txtSpellName.Text
    Spell(EditorIndex).ClassReq = frmEditor.cmbSpellClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmEditor.scrlSpellLevelReq.Value
    Spell(EditorIndex).Type = frmEditor.cmbSpellType.ListIndex
    If Spell(EditorIndex).Type <> SPELL_TYPE_GIVEITEM Then
        Spell(EditorIndex).Data1 = frmEditor.scrlSpellVitalMod.Value
    Else
        Spell(EditorIndex).Data1 = frmEditor.scrlSpellItemNum.Value
        Spell(EditorIndex).Data2 = frmEditor.scrlSpellItemValue.Value
    End If
    Spell(EditorIndex).Data3 = 0
    
    Call SendSaveSpell(EditorIndex)
End Sub

Public Sub ClassEditorInit()
    frmEditor.txtClassName.Text = Trim$(Class(EditorIndex - 1).Name)
    frmEditor.scrlClassSprite.Value = Class(EditorIndex - 1).Sprite
    frmEditor.txtClassHP.Text = CStr(Class(EditorIndex - 1).HP)
    frmEditor.txtClassMP.Text = CStr(Class(EditorIndex - 1).MP)
    frmEditor.txtClassSP.Text = CStr(Class(EditorIndex - 1).SP)
    frmEditor.txtClassSTR.Text = CStr(Class(EditorIndex - 1).STR)
    frmEditor.txtClassDEF.Text = CStr(Class(EditorIndex - 1).DEF)
    frmEditor.txtClassMAGI.Text = CStr(Class(EditorIndex - 1).MAGI)
    frmEditor.txtClassSPD.Text = CStr(Class(EditorIndex - 1).Speed)
    frmEditor.txtClassMap.Text = CStr(Class(EditorIndex - 1).Map)
    frmEditor.txtClassX.Text = CStr(Class(EditorIndex - 1).X)
    frmEditor.txtClassY.Text = CStr(Class(EditorIndex - 1).Y)
End Sub

Public Sub ClassEditorSave()
    Class(EditorIndex - 1).Name = Trim$(frmEditor.txtClassName.Text)
    Class(EditorIndex - 1).Sprite = frmEditor.scrlClassSprite.Value
    Class(EditorIndex - 1).HP = Val(frmEditor.txtClassHP.Text)
    Class(EditorIndex - 1).MP = Val(frmEditor.txtClassMP.Text)
    Class(EditorIndex - 1).SP = Val(frmEditor.txtClassSP.Text)
    Class(EditorIndex - 1).STR = Val(frmEditor.txtClassSTR.Text)
    Class(EditorIndex - 1).DEF = Val(frmEditor.txtClassDEF.Text)
    Class(EditorIndex - 1).MAGI = Val(frmEditor.txtClassMAGI.Text)
    Class(EditorIndex - 1).Speed = Val(frmEditor.txtClassSPD.Text)
    Class(EditorIndex - 1).Map = Val(frmEditor.txtClassMap.Text)
    Class(EditorIndex - 1).X = Val(frmEditor.txtClassX.Text)
    Class(EditorIndex - 1).Y = Val(frmEditor.txtClassY.Text)
    Call SendSaveClass(EditorIndex - 1)
End Sub

Public Sub UpdateInventory()
    Call DrawInventory
End Sub

Sub ResizeGUI()
    If frmDualSolace.WindowState <> vbMinimized Then
        frmDualSolace.txtChat.Height = Int(frmDualSolace.Height / Screen.TwipsPerPixelY) - frmDualSolace.txtChat.top - 32
        frmDualSolace.txtChat.Width = Int(frmDualSolace.Width / Screen.TwipsPerPixelX) - 8
    End If
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim x1 As Long, y1 As Long

    x1 = Int(X / PIC_X)
    y1 = Int(Y / PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub ClearTempTile()
Dim X As Long, Y As Long

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            TempTile(X, Y).DoorOpen = NO
        Next X
    Next Y
End Sub

Sub ClearPlayer(ByVal Index As Long)
Dim i As Long
Dim n As Long

    Player(Index).Name = ""
    Player(Index).Class = 0
    Player(Index).Level = 0
    Player(Index).Sprite = 0
    Player(Index).EXP = 0
    Player(Index).Access = 0
    Player(Index).PK = NO
        
    Player(Index).HP = 0
    Player(Index).MP = 0
    Player(Index).SP = 0
        
    Player(Index).STR = 0
    Player(Index).DEF = 0
    Player(Index).Speed = 0
    Player(Index).MAGI = 0
        
    For n = 1 To MAX_FRIENDS
        Player(Index).Friends(n) = ""
    Next n
        
    For n = 1 To MAX_INV
        Player(Index).Inv(n).num = 0
        Player(Index).Inv(n).Value = 0
        Player(Index).Inv(n).Dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
        
    Player(Index).Map = 0
    Player(Index).X = 0
    Player(Index).Y = 0
    Player(Index).Dir = 0
    
    ' Client use only
    Player(Index).MaxHP = 0
    Player(Index).MaxMP = 0
    Player(Index).MaxSP = 0
    Player(Index).XOffset = 0
    Player(Index).YOffset = 0
    Player(Index).Moving = 0
    Player(Index).Attacking = 0
    Player(Index).AttackTimer = 0
    Player(Index).MapGetTimer = 0
    Player(Index).CastedSpell = NO
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = ""
    
    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
End Sub

Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next i
End Sub

Sub ClearMapItem(ByVal Index As Long)
    MapItem(Index).num = 0
    MapItem(Index).Value = 0
    MapItem(Index).Dur = 0
    MapItem(Index).X = 0
    MapItem(Index).Y = 0
End Sub

Sub ClearMap()
Dim i As Long
Dim X As Long
Dim Y As Long

    Map.Name = ""
    Map.Revision = 0
    Map.Moral = 0
    Map.Up = 0
    Map.Down = 0
    Map.Left = 0
    Map.Right = 0
        
    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            With Map.Tile(X, Y)
                .Ground = 0
                .Mask = 0
                .Mask2 = 0
                .Anim = 0
                .Anim2 = 0
                .Fringe = 0
                .FringeAnim = 0
                .Fringe2 = 0
                .Walkable = True
                .Blocked = False
                .North = False
                .West = False
                .East = False
                .South = False
                .Warp = False
                .WarpMap = 0
                .WarpX = 0
                .WarpY = 0
                .Item = False
                .ItemNum = 0
                .ItemValue = 0
                .NpcAvoid = False
                .Key = False
                .KeyNum = 0
                .KeyTake = 0
                .KeyOpen = False
                .KeyOpenX = 0
                .KeyOpenY = 0
                .Bank = False
                .Shop = False
                .ShopNum = 0
                .Heal = False
                .HealValue = 0
                .Damage = False
                .DamageValue = 0
            End With
        Next X
    Next Y
End Sub

Sub ClearMapItems()
Dim X As Long

    For X = 1 To MAX_MAP_ITEMS
        Call ClearMapItem(X)
    Next X
End Sub

Sub ClearMapNpc(ByVal Index As Long)
    MapNpc(Index).num = 0
    MapNpc(Index).Target = 0
    MapNpc(Index).HP = 0
    MapNpc(Index).MaxHP = 0
    MapNpc(Index).MP = 0
    MapNpc(Index).SP = 0
    MapNpc(Index).Map = 0
    MapNpc(Index).X = 0
    MapNpc(Index).Y = 0
    MapNpc(Index).Dir = 0
    
    ' Client use only
    MapNpc(Index).XOffset = 0
    MapNpc(Index).YOffset = 0
    MapNpc(Index).Moving = 0
    MapNpc(Index).Attacking = 0
    MapNpc(Index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call ClearMapNpc(i)
    Next i
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    If Index > 0 Then
        GetPlayerName = Trim$(Player(Index).Name)
    Else
        GetPlayerName = vbNullString
    End If
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).EXP
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal EXP As Long)
    Player(Index).EXP = EXP
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (GetPlayerLevel(Index) + 1) * (GetPlayerSTR(Index) + GetPlayerDEF(Index) + GetPlayerMAGI(Index) + GetPlayerSPEED(Index)) * 25
End Function

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).HP = HP
    
    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).HP = GetPlayerMaxHP(Index)
    End If
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).MP = GetPlayerMaxMP(Index)
    End If
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).SP = GetPlayerMaxSP(Index)
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    GetPlayerMaxHP = Player(Index).MaxHP
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    GetPlayerMaxMP = Player(Index).MaxMP
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    GetPlayerMaxSP = Player(Index).MaxSP
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    GetPlayerSTR = Player(Index).STR
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    GetPlayerDEF = Player(Index).DEF
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    GetPlayerSPEED = Player(Index).Speed
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).MAGI
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
    Player(Index).MAGI = MAGI
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).BankInv(InvSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).BankInv(InvSlot).num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).BankInv(InvSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).BankInv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).BankInv(InvSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).BankInv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ShieldSlot = InvNum
End Sub

Function IsFriend(ByVal Index As Long, ByVal Name As String) As Boolean
Dim n As Integer
IsFriend = False
For n = 1 To MAX_FRIENDS
    If Player(Index).Friends(n) = Name Then
        IsFriend = True
        Exit Function
    End If
Next n
End Function

'Function SetMouseCursor(CursorType As Long)
'Dim hCursor As Long
'hCursor = LoadCursorLong(0&, CursorType)
'hCursor = SetCursor(hCursor)
'End Function

Public Sub SetFormSize(psPath As String, Form As Form)
  Dim f As Integer
  Dim tmp As String
  Dim FileHeader As BITMAPFILEHEADER
  Dim InfoHeader As BITMAPINFOHEADER
  
  On Error Resume Next
 
  f = FreeFile
  Open psPath For Binary Access Read As #f
  Get #f, , FileHeader
  Get #f, , InfoHeader
  Close #f

  ' Set to pixel mode
  Form.ScaleMode = 3

  Form.ScaleWidth = InfoHeader.biWidth
  Form.ScaleHeight = InfoHeader.biHeight
  'Form.Width = Form.ScaleX(InfoHeader.biWidth, vbPixels, vbTwips)
  'Form.Height = Form.ScaleY(InfoHeader.biHeight, vbPixels, vbTwips)
  
  ' Set to twips mode
  Form.ScaleMode = 1
End Sub

Public Sub SetPicSize(psPath As String, Pic As PictureBox, Optional ByVal Pixels As Boolean = False)
  Dim f As Integer
  Dim tmp As String
  Dim FileHeader As BITMAPFILEHEADER
  Dim InfoHeader As BITMAPINFOHEADER
  
  On Error Resume Next
 
  f = FreeFile
  Open psPath For Binary Access Read As #f
  Get #f, , FileHeader
  Get #f, , InfoHeader
  Close #f

If Pixels = False Then
  Pic.Width = InfoHeader.biWidth
  Pic.Height = InfoHeader.biHeight
Else
  Pic.ScaleMode = 3
  Pic.ScaleWidth = InfoHeader.biWidth
  Pic.ScaleHeight = InfoHeader.biHeight
  Pic.ScaleMode = 1
End If
End Sub

