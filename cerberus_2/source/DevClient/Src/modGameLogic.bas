Attribute VB_Name = "modGameLogic"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA
 
Option Explicit

Public Sub Main()
Dim i As Long

    ' Check for configuration file
    If FileExist("config.ini", False) = False Then
        PutVar App.Path & "\config.ini", "IPCONFIG", "IP", "127.0.0.1"
        PutVar App.Path & "\config.ini", "IPCONFIG", "Always", 0
        PutVar App.Path & "\config.ini", "IPCONFIG", "Port", 7000
        PutVar App.Path & "\config.ini", "CONFIG", "GameName", "Cerberus Default"
        PutVar App.Path & "\config.ini", "CONFIG", "WebSite", "http://www.webaddress.com"
        PutVar App.Path & "\config.ini", "CONFIG", "Account", "Name"
        PutVar App.Path & "\config.ini", "CONFIG", "Password", "Password"
    End If

    'Check for IP configuration
    If GetVar(App.Path & "\config.ini", "IPCONFIG", "Always") = NO Then
        frmIPConfig.Show vbModal
    Else
        GAME_IP = GetVar(App.Path & "\config.ini", "IPCONFIG", "IP")
        GAME_PORT = GetVar(App.Path & "\config.ini", "IPCONFIG", "Port")
    End If

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
    InSpellEditor = False
    InSkillEditor = False
    InShopEditor = False
    InQuestEditor = False
    InGUIEditor = False
    
    ' Clear out players
    'For i = 1 To MAX_PLAYERS
        'Call ClearPlayer(i)
    'Next i
    Call ClearTempTile
    Call ClearPushTile
    
    frmSendGetData.Visible = True
    Call SetStatus("Initializing TCP settings...")
    Call TcpInit
    frmMainMenu.Visible = True
    frmSendGetData.Visible = False
End Sub

Sub SetStatus(ByVal Caption As String)
    frmSendGetData.lblStatus.Caption = Caption
End Sub

Sub GameDestroy()
    Call DestroyDirectX
    End
End Sub

Public Sub MenuState(ByVal State As Long)
    frmSendGetData.Visible = True
    Call SetStatus("Connecting to server...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmMainMenu.picNewAccountMenu.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMainMenu.txtNameNew.Text, frmMainMenu.txtPasswordNew.Text)
            End If
            
        Case MENU_STATE_DELACCOUNT
            frmMainMenu.picDeleteAccountMenu.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending account deletion request ...")
                Call SendDelAccount(frmMainMenu.txtNameDelete.Text, frmMainMenu.txtPasswordDelete.Text)
            End If
        
        Case MENU_STATE_LOGIN
            frmMainMenu.picLoginMenu.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMainMenu.txtNameLogin, frmMainMenu.txtPasswordLogin.Text)
            End If
        
        Case MENU_STATE_NEWCHAR
            frmMainMenu.picChars.Visible = False
            frmMainMenu.txtCharName.Text = ""
            frmMainMenu.optMale.Value = True
            Call SetStatus("Connected, getting available classes...")
            Call SendGetClasses
            
        Case MENU_STATE_ADDCHAR
            frmMainMenu.picCreateChar.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character addition data...")
                If frmMainMenu.optMale.Value = True Then
                    Call SendAddChar(frmMainMenu.txtCharName.Text, 0, frmMainMenu.cmbClass.ListIndex, frmMainMenu.lstChars.ListIndex + 1)
                Else
                    Call SendAddChar(frmMainMenu.txtCharName.Text, 1, frmMainMenu.cmbClass.ListIndex, frmMainMenu.lstChars.ListIndex + 1)
                End If
            End If
        
        Case MENU_STATE_DELCHAR
            frmMainMenu.picChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character deletion request...")
                Call SendDelChar(frmMainMenu.lstChars.ListIndex + 1)
            End If
            
        Case MENU_STATE_USECHAR
            frmMainMenu.picChars.Visible = False
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending char data...")
                Call SendUseChar(frmMainMenu.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected Then
        frmMainMenu.Visible = True
        frmSendGetData.Visible = False
        Call MsgBox("Sorry, the server seems to be down.  Please try to reconnect in a few minutes or visit " & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub

Sub GameInit()
    frmCClient.Visible = True
    frmMainMenu.Visible = False
    frmSendGetData.Visible = False
    'Call ResizeGUI
    Call InitDirectX
End Sub

Public Sub GameLoop()

Dim Tick As Long
Dim TickFPS As Long
Dim FPS As Long
Dim x As Long
Dim y As Long
Dim i As Long
Dim rec_back As RECT
    
    ' Set the focus
    frmCClient.picScreen.SetFocus
    
    ' Set font
    Call SetFont(FONT_NAME, FONT_SIZE)
    
    '' Used for calculating fps
    'TickFPS = GetTickCount
    'FPS = 0
    
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
                
        ' Blit out tiles layers ground/mask/anim
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltTile(x, y)
            Next x
        Next y
        
        ' Blit out pushblock movement
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If PushTile(x, y).Moving > 0 Then
                    Call BltPushBlock(x, y)
                End If
            Next x
        Next y
                    
        ' Blit out the items
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).Num > 0 Then
                Call BltItem(i)
            End If
        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpc(i)
        Next i
        
        ' Blit out the resources
        For i = 1 To MAX_MAP_RESOURCES
            Call BltResource(i)
        Next i
        
        ' Blit out players
        For i = 1 To HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayer(i)
            End If
        Next i
        
        ' Blit out the npcs tops
        For i = 1 To MAX_MAP_NPCS
            Call BltNpcTop(i)
        Next i
        
        ' Blit out players tops
        For i = 1 To HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call BltPlayerTop(i)
            End If
        Next i
        
        ' Blit out the resources tops
        For i = 1 To MAX_MAP_RESOURCES
            Call BltResourceTop(i)
        Next i
                
        ' Blit out tile layer fringe
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltFringeTile(x, y)
            Next x
        Next y
        
        ' Blit out tile layer light
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Call BltLightTile(x, y)
            Next x
        Next y
                
        ' Lock the backbuffer so we can draw text and names
        TexthDC = DD_BackBuffer.GetDC
        
        'For i = 1 To HighIndex
            'If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                'Call BltPlayerName(i)
            'End If
        'Next i
                
        ' Blit out attribs if in editor
        If InEditor Then
            If frmCClient.optDirectionView.Value = False And frmCClient.optLayers.Value = False And frmCClient.optBuildLayer.Value = False Then
                For y = 0 To MAX_MAPY
                    For x = 0 To MAX_MAPX
                        With Map.Tile(x, y)
                            If .Type = TILE_TYPE_BLOCKED Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "B", QBColor(BrightRed))
                            If .Type = TILE_TYPE_WARP Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "W", QBColor(BrightBlue))
                            If .Type = TILE_TYPE_ITEM Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "I", QBColor(White))
                            If .Type = TILE_TYPE_NPCAVOID Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "N", QBColor(White))
                            If .Type = TILE_TYPE_KEY Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "K", QBColor(White))
                            If .Type = TILE_TYPE_KEYOPEN Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "O", QBColor(White))
                            If .Type = TILE_TYPE_PUSHBLOCK Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "P", QBColor(White))
                            If .Type = TILE_TYPE_NSPAWN Then Call DrawText(TexthDC, x * PIC_X + 4, y * PIC_Y + 4, "Ns", QBColor(Green))
                            If .Type = TILE_TYPE_RSPAWN Then Call DrawText(TexthDC, x * PIC_X + 4, y * PIC_Y + 4, "Rs", QBColor(Green))
                     End With
                    Next x
                Next y
            ElseIf frmCClient.optLayers.Value = False And frmCClient.optAttribs.Value = False And frmCClient.optDirectionView.Value = False Then
                For y = 0 To MAX_MAPY
                    For x = 0 To MAX_MAPX
                        With Map.Tile(x, y)
                            If .Build = 1 Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "L", QBColor(Grey))
                     End With
                    Next x
                Next y
            End If
        End If
        
        '' Blit the text they are putting in
        'Call DrawText(TexthDC, 0, (MAX_MAPY + 1) * PIC_Y - 20, MyText, RGB(255, 255, 255))
        
        ' Add the text they enter to the chat input box
        frmCClient.lblChat.Caption = MyText
        
        ' Draw map name
        If Map.Moral = MAP_MORAL_NONE Then
            Call DrawText(TexthDC, Int((MAX_MAPX + 1) * PIC_X / 2) - (Int(Len(Trim(Map.Name)) / 2) * 8), 1, Trim(Map.Name), QBColor(BrightRed))
        Else
            Call DrawText(TexthDC, Int((MAX_MAPX + 1) * PIC_X / 2) - (Int(Len(Trim(Map.Name)) / 2) * 8), 1, Trim(Map.Name), QBColor(White))
        End If
        
        ' Check if we are getting a map, and if we are tell them so
        If GettingMap = True Then
            Call DrawText(TexthDC, 50, 50, "Receiving Map...", QBColor(BrightCyan))
        End If
                        
        ' Release DC
        Call DD_BackBuffer.ReleaseDC(TexthDC)
        
        ' Used to blit out directional blocking arrows
        If InEditor Then
             If frmCClient.optDirectionView.Value = True Then
                 Call BltDirectionArrows
             End If
        End If
        
        ' Get the rect for the back buffer to blit from
        With rec
            .Top = 0
            .Bottom = (MAX_MAPY + 1) * PIC_Y
            .Left = 0
            .Right = (MAX_MAPX + 1) * PIC_X
        End With
        
        ' Get the rect to blit to
        Call DX.GetWindowRect(frmCClient.picScreen.hWnd, rec_pos)
        With rec_pos
            .Bottom = .Top + ((MAX_MAPY + 1) * PIC_Y)
            .Right = .Left + ((MAX_MAPX + 1) * PIC_X)
        End With
        
        ' Blit the backbuffer
        Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec, DDBLT_WAIT)
        
        ' Check if player is trying to move
        Call CheckMovement
        
        '' Check to see if player is trying to attack
        'Call CheckAttack
        
        ' Process player movements (actually move them)
        For i = 1 To HighIndex
            If IsPlaying(i) Then
                Call ProcessMovement(i)
            End If
        Next i
        
        ' Process pushblock movements (actually move them)
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                If PushTile(x, y).Moving > 0 Then
                    Call ProcessPushBlock(x, y)
                End If
            Next x
        Next y
        
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
                
        '' Lock fps
        'Do While GetTickCount < Tick + 50
            'DoEvents
        'Loop
        
        '' Calculate fps
        'If GetTickCount > TickFPS + 1000 Then
            'GameFPS = FPS
            'TickFPS = GetTickCount
            'FPS = 0
        'Else
            'FPS = FPS + 1
        'End If
        
        DoEvents
        
    Loop
    
    frmCClient.Visible = False
    frmSendGetData.Visible = True
    Call SetStatus("Destroying game data...")
    
    ' Shutdown the game
    Call GameDestroy
    
    ' Report disconnection if server disconnects
    If IsConnected = False Then
        Call MsgBox("Thank you for playing " & GAME_NAME & "!", vbOKOnly, GAME_NAME)
    End If
End Sub

Public Sub BltTile(ByVal x As Long, ByVal y As Long)
Dim Ground As Long
Dim Mask As Long
Dim Mask2 As Long
Dim Anim As Long

    With Map.Tile(x, y)
        Ground = .Ground
        Mask = .Mask
        Mask2 = .Mask2
        Anim = .Anim
    End With
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    With rec
        .Top = Int(Ground / 14) * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = (Ground - Int(Ground / 14) * 14) * PIC_X
        .Right = .Left + PIC_X
    End With
    Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT)
    
    If Mask > 0 And TempTile(x, y).DoorOpen = NO And PushTile(x, y).Pushed = NO And PushTile(x, y).Moving = 0 Then
        With rec
            .Top = Int(Mask / 14) * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = (Mask - Int(Mask / 14) * 14) * PIC_X
            .Right = .Left + PIC_X
        End With
        Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
    
    'If (MapAnim = 0) Or (Anim <= 0) Then
        '' Is there an animation tile to plot?
        'If Mask > 0 And TempTile(x, y).DoorOpen = NO Then
            'With rec
                '.top = Int(Mask / 14) * PIC_Y
                '.Bottom = .top + PIC_Y
                '.Left = (Mask - Int(Mask / 14) * 14) * PIC_X
                '.Right = .Left + PIC_X
            'End With
            'Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        'End If
    'Else
        '' Is there a second animation tile to plot?
        'If Anim > 0 Then
            'With rec
                '.top = Int(Anim / 14) * PIC_Y
                '.Bottom = .top + PIC_Y
                '.Left = (Anim - Int(Anim / 14) * 14) * PIC_X
                '.Right = .Left + PIC_X
            'End With
            'Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        'End If
    'End If
    
    If (MapAnim = 0) Or (Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask2 > 0 Then
             rec.Top = Int(Mask2 / 14) * PIC_Y
             rec.Bottom = rec.Top + PIC_Y
             rec.Left = (Mask2 - Int(Mask2 / 14) * 14) * PIC_X
             rec.Right = rec.Left + PIC_X
             'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
             Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim > 0 Then
             rec.Top = Int(Anim / 14) * PIC_Y
             rec.Bottom = rec.Top + PIC_Y
             rec.Left = (Anim - Int(Anim / 14) * 14) * PIC_X
             rec.Right = rec.Left + PIC_X
             'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
             Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Public Sub BltPushBlock(ByVal x As Integer, ByVal y As Integer)
Dim x1 As Long, y1 As Long

    ' Only used if ever want to switch to blt rather then bltfast
    'With rec_pos
        '.top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
        '.Bottom = .top + PIC_Y
        '.Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
        '.Right = .Left + PIC_X
    'End With
    
    With rec
        .Top = Int(Map.Tile(x, y).Mask / 14) * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = (Map.Tile(x, y).Mask - Int(Map.Tile(x, y).Mask / 14) * 14) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    x1 = x * PIC_X + PushTile(x, y).XOffset
    y1 = y * PIC_Y + PushTile(x, y).YOffset
    
    ' Check if its out of bounds because of the offset
    If y1 < 0 Then
        y1 = 0
        With rec
            .Top = .Top + (y1 * -1)
        End With
    End If
        
    Call DD_BackBuffer.BltFast(x1, y1, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltItem(ByVal ItemNum As Long)

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapItem(ItemNum).y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = MapItem(ItemNum).x * PIC_X
        .Right = .Left + PIC_X
    End With

    With rec
        .Top = Item(MapItem(ItemNum).Num).Pic * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With
    
    Call DD_BackBuffer.BltFast(MapItem(ItemNum).x * PIC_X, MapItem(ItemNum).y * PIC_Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltFringeTile(ByVal x As Long, ByVal y As Long)
Dim Fringe As Long
Dim Fringe2 As Long
Dim FAnim As Long

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Fringe = Map.Tile(x, y).Fringe
    Fringe2 = Map.Tile(x, y).Fringe2
    FAnim = Map.Tile(x, y).FAnim
        
    If Fringe > 0 Then
        rec.Top = Int(Fringe / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (Fringe - Int(Fringe / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

    If (MapAnim = 0) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe2 > 0 Then
            rec.Top = Int(Fringe2 / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Fringe2 - Int(Fringe2 / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If FAnim > 0 Then
            rec.Top = Int(FAnim / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (FAnim - Int(FAnim / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Public Sub BltLightTile(ByVal x As Long, ByVal y As Long)
Dim Light As Long

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Light = Map.Tile(x, y).Light
        
    If Light > 0 Then
        With rec
            .Top = Int(Light / 14) * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = (Light - Int(Light / 14) * 14) * PIC_X
            .Right = .Left + PIC_X
        End With
        Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Public Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long

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
               Anim = 0
               If (Player(Index).YOffset < PIC_Y / 3) Then
               Anim = 1
               ElseIf (Player(Index).YOffset > PIC_Y / 3) And ((Player(Index).YOffset > PIC_Y / 3 * 2)) Then
               Anim = 2
               End If
           Case DIR_DOWN
               Anim = 1
               If (Player(Index).YOffset < PIC_X / 4 * -1) Then Anim = 0
               If (Player(Index).YOffset < PIC_X / 2 * -1) Then Anim = 2
           Case DIR_LEFT
               Anim = 0
               If (Player(Index).XOffset < PIC_Y / 3) Then
               Anim = 1
               ElseIf (Player(Index).XOffset > PIC_Y / 3) And ((Player(Index).XOffset > PIC_Y / 3 * 2)) Then
               Anim = 2
               End If
           Case DIR_RIGHT
               Anim = 0
               If (Player(Index).XOffset < PIC_Y / 4 * -1) Then Anim = 1
               If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = 2
       End Select
    Else
        If Player(Index).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    '' Check to see if we want to stop making him attack
    'With Player(Index)
        'If .AttackTimer + 1000 < GetTickCount Then
            '.Attacking = 0
            '.AttackTimer = 0
        'End If
    'End With
    
    With rec
        .Top = (((GetPlayerSprite(Index) * 2) * PIC_Y) + PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        With rec
            .Top = .Top + (y * -1)
        End With
    End If
    
    '' Check if its out of bounds because of the offset
    'y = y - 32
    'If y < 0 And y > -32 Then
        'With rec
            '.top = .top - y
            'y = 0
        'End With
    'End If
        
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltPlayerTop(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long

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
               Anim = 0
               If (Player(Index).YOffset < PIC_Y / 3) Then
               Anim = 1
               ElseIf (Player(Index).YOffset > PIC_Y / 3) And ((Player(Index).YOffset > PIC_Y / 3 * 2)) Then
               Anim = 2
               End If
           Case DIR_DOWN
               Anim = 1
               If (Player(Index).YOffset < PIC_X / 4 * -1) Then Anim = 0
               If (Player(Index).YOffset < PIC_X / 2 * -1) Then Anim = 2
           Case DIR_LEFT
               Anim = 0
               If (Player(Index).XOffset < PIC_Y / 3) Then
               Anim = 1
               ElseIf (Player(Index).XOffset > PIC_Y / 3) And ((Player(Index).XOffset > PIC_Y / 3 * 2)) Then
               Anim = 2
               End If
           Case DIR_RIGHT
               Anim = 0
               If (Player(Index).XOffset < PIC_Y / 4 * -1) Then Anim = 1
               If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = 2
       End Select
    Else
        If Player(Index).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    '' Check to see if we want to stop making him attack
    'With Player(Index)
        'If .AttackTimer + 1000 < GetTickCount Then
            '.Attacking = 0
            '.AttackTimer = 0
        'End If
    'End With
    
    With rec
        .Top = ((GetPlayerSprite(Index) * 2) * PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
    
    '' Check if its out of bounds because of the offset
    'If y < 0 Then
        'y = 0
        'With rec
            '.top = .top + (y * -1)
        'End With
    'End If
    
    ' Check if its out of bounds because of the offset
    y = y - 32
    If y < 0 And y > -32 Then
        With rec
            .Top = .Top - y
            y = 0
        End With
    End If
        
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltDirectionArrows()
Dim x As Long, y As Long
Dim Blocked As Boolean ' If tile has at least one dir, change center to circle

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
             With Map.Tile(x, y)
            
                 Blocked = False ' Reset for a new tile
                
                 ' Lets check what to Blt for the Up arrow
                 If .WalkUp = 1 Then
                     With rec
                          .Top = 0
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 3
                          .Right = .Left + PIC_Y
                     End With
                     Blocked = True
                 Else
                     With rec
                          .Top = PIC_X
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 3
                          .Right = .Left + PIC_Y
                     End With
                 End If
                 Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_DirectionSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                 ' Now lets check For Down Arrow
                 If .WalkDown = 1 Then
                     With rec
                          .Top = 0
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 2
                          .Right = .Left + PIC_Y
                     End With
                     Blocked = True
                 Else
                     With rec
                          .Top = PIC_X
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 2
                          .Right = .Left + PIC_Y
                     End With
                 End If
                 Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_DirectionSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                 ' Now lets check For Left Arrow
                 If .WalkLeft = 1 Then
                     With rec
                          .Top = 0
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 0
                          .Right = .Left + PIC_Y
                     End With
                     Blocked = True
                 Else
                     With rec
                          .Top = PIC_X
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 0
                          .Right = .Left + PIC_Y
                     End With
                 End If
                 Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_DirectionSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                 ' Now lets check For Right Arrow
                 If .WalkRight = 1 Then
                     With rec
                          .Top = 0
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 1
                          .Right = .Left + PIC_Y
                     End With
                     Blocked = True
                 Else
                     With rec
                          .Top = PIC_X
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 1
                          .Right = .Left + PIC_Y
                     End With
                 End If
                 Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_DirectionSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                 ' If tile is totally blocked, place cross in center
                 If Blocked Then
                     With rec
                          .Top = 0
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 4
                          .Right = .Left + PIC_Y
                     End With
                 Else
                     With rec
                          .Top = PIC_X
                          .Bottom = .Top + PIC_X
                          .Left = PIC_Y * 4
                          .Right = .Left + PIC_Y
                     End With
                 End If
                 Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_DirectionSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                
                 ' We need to draw a grid so we can tell which tiles which (First set pen width to 1)
                 DD_BackBuffer.DrawLine x * PIC_X, y * PIC_Y, x * PIC_X, y * PIC_Y + (PIC_Y * MAX_MAPY) + 1
                 DD_BackBuffer.DrawLine x * PIC_X, y * PIC_Y, x * PIC_X + (PIC_X * MAX_MAPX) + 1, y * PIC_Y
                              End With
        Next x
        
        DoEvents
    Next y
End Sub

Public Sub BltNpc(ByVal MapNpcNum As Long)
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
    
    '' Check to see if we want to stop making him attack
    'With MapNpc(MapNpcNum)
        'If .AttackTimer + 1000 < GetTickCount Then
            '.Attacking = 0
            '.AttackTimer = 0
        'End If
    'End With
    
    With rec
        .Top = (((Npc(MapNpc(MapNpcNum).Num).Sprite * 2) * PIC_Y) + PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    With MapNpc(MapNpcNum)
        x = .x * PIC_X + .XOffset
        y = .y * PIC_Y + .YOffset - 4
    End With
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        With rec
            .Top = .Top + (y * -1)
        End With
    End If
    
    '' Check if its out of bounds because of the offset
    'y = y - 32
    'If y < 0 And y > -32 Then
        'With rec
                '.top = .top - y
                'y = 0
        'End With
    'End If
        
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltNpcTop(ByVal MapNpcNum As Long)
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
    
    '' Check to see if we want to stop making him attack
    'With MapNpc(MapNpcNum)
        'If .AttackTimer + 1000 < GetTickCount Then
            '.Attacking = 0
            '.AttackTimer = 0
        'End If
    'End With
    
    With rec
        .Top = ((Npc(MapNpc(MapNpcNum).Num).Sprite * 2) * PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    With MapNpc(MapNpcNum)
        x = .x * PIC_X + .XOffset
        y = .y * PIC_Y + .YOffset - 4
    End With
    
    '' Check if its out of bounds because of the offset
    'If y < 0 Then
        'y = 0
        'With rec
            '.top = .top + (y * -1)
        'End With
    'End If
    
    ' Check if its out of bounds because of the offset
    y = y - 32
    If y < 0 And y > -32 Then
        With rec
            .Top = .Top - y
            y = 0
        End With
    End If
        
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltResource(ByVal MapResourceNum As Long)
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapResource(MapResourceNum).Num <= 0 Then
        Exit Sub
    End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapResource(MapResourceNum).y * PIC_Y '+ MapResource(MapResourceNum).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapResource(MapResourceNum).x * PIC_X '+ MapResource(MapResourceNum).XOffset
        .Right = .Left + PIC_X
    End With
    
    '' Check for animation
    'Anim = 0
    'If MapNpc(MapNpcNum).Attacking = 0 Then
        'Select Case MapNpc(MapNpcNum).Dir
            'Case DIR_UP
                'If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            'Case DIR_DOWN
                'If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            'Case DIR_LEFT
                'If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            'Case DIR_RIGHT
                'If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        'End Select
    'Else
        'If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            'Anim = 2
        'End If
    'End If
    
    '' Check to see if we want to stop making him attack
    'With MapNpc(MapNpcNum)
        'If .AttackTimer + 1000 < GetTickCount Then
            '.Attacking = 0
            '.AttackTimer = 0
        'End If
    'End With
    
    If Npc(MapResource(MapResourceNum).Num).Big = 1 Then
        With rec
            .Top = (((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y) + PIC_Y)
            .Bottom = .Top + PIC_Y
            .Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + 64
        End With
    ElseIf Npc(MapResource(MapResourceNum).Num).Big = 2 Then
        With rec
            .Top = (((Npc(MapResource(MapResourceNum).Num).Sprite * 4) * PIC_Y) + (3 * PIC_Y))
            .Bottom = .Top + PIC_Y
            .Left = PIC_X '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + PIC_X
        End With
    'ElseIf Npc(MapResource(MapResourceNum).Num).Big = 3 Then
        'With rec
            '.top = (((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y) + PIC_Y)
            '.Bottom = .top + PIC_Y
            '.Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            '.Right = .Left + PIC_X
        'End With
    Else
        With rec
        .Top = (((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y) + PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
        End With
    End If
    
    With MapResource(MapResourceNum)
        x = .x * PIC_X '+ .XOffset
        y = .y * PIC_Y '+ .YOffset - 4
    End With
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        With rec
            .Top = .Top + (y * -1)
        End With
    End If
    
    '' Check if its out of bounds because of the offset
    'y = y - 32
    'If y < 0 And y > -32 Then
        'With rec
                '.top = .top - y
                'y = 0
        'End With
    'End If
    If Npc(MapResource(MapResourceNum).Num).Big = 1 Then
        Call DD_BackBuffer.BltFast(x, y, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    ElseIf Npc(MapResource(MapResourceNum).Num).Big = 2 Then
        Call DD_BackBuffer.BltFast(x, y, DD_TreeSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    'ElseIf Npc(MapResource)(MapResourceNum).Num).Big = 3 Then
        'Call DD_BackBuffer.BltFast(x, y, DD_BuildingSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Public Sub BltResourceTop(ByVal MapResourceNum As Long)
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapResource(MapResourceNum).Num <= 0 Then
        Exit Sub
    End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapResource(MapResourceNum).y * PIC_Y '+ MapResource(MapResourceNum).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapResource(MapResourceNum).x * PIC_X '+ MapResource(MapResourceNum).XOffset
        .Right = .Left + PIC_X
    End With
    
    '' Check for animation
    'Anim = 0
    'If MapNpc(MapNpcNum).Attacking = 0 Then
        'Select Case MapNpc(MapNpcNum).Dir
            'Case DIR_UP
                'If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            'Case DIR_DOWN
                'If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            'Case DIR_LEFT
                'If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            'Case DIR_RIGHT
                'If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        'End Select
    'Else
        'If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            'Anim = 2
        'End If
    'End If
    
    '' Check to see if we want to stop making him attack
    'With MapNpc(MapNpcNum)
        'If .AttackTimer + 1000 < GetTickCount Then
            '.Attacking = 0
            '.AttackTimer = 0
        'End If
    'End With
    
    If Npc(MapResource(MapResourceNum).Num).Big = 1 Then
        With rec
            .Top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y)
            .Bottom = .Top + PIC_Y
            .Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + 64
        End With
    ElseIf Npc(MapResource(MapResourceNum).Num).Big = 2 Then
        With rec
            .Top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 4) * PIC_Y)
            .Bottom = .Top + (4 * PIC_Y)
            .Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + PIC_X
        End With
        With rec1
            .Top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 4) * PIC_Y)
            .Bottom = .Top + (3 * PIC_Y)
            .Left = PIC_X '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + PIC_X
        End With
        With rec2
            .Top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 4) * PIC_Y)
            .Bottom = .Top + (4 * PIC_Y)
            .Left = 2 * PIC_X '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + PIC_X
        End With
    'ElseIf Npc(MapResource(MapResourceNum).Num).Big = 3 Then
        'With rec
            '.top = (((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y)
            '.Bottom = .top + PIC_Y
            '.Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            '.Right = .Left + PIC_X
        'End With
    Else
        With rec
        .Top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y)
        .Bottom = .Top + PIC_Y
        .Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
        End With
    End If
    
    With MapResource(MapResourceNum)
        x = .x * PIC_X '+ .XOffset
        y = .y * PIC_Y '+ .YOffset - 4
    End With
    
    '' Check if its out of bounds because of the offset
    'If y < 0 Then
        'y = 0
        'With rec
            '.top = .top + (y * -1)
        'End With
    'End If
    
    ' Check if its out of bounds because of the offset
    y = y - 32
    If y < 0 And y > -32 Then
        With rec
            .Top = .Top - y
            y = 0
        End With
    End If
        
    If Npc(MapResource(MapResourceNum).Num).Big = 1 Then
        Call DD_BackBuffer.BltFast(x, y, DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    ElseIf Npc(MapResource(MapResourceNum).Num).Big = 2 Then
        Call DD_BackBuffer.BltFast(x - 32, y - 64, DD_TreeSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Call DD_BackBuffer.BltFast(x, y - 64, DD_TreeSpriteSurf, rec1, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Call DD_BackBuffer.BltFast(x + 32, y - 64, DD_TreeSpriteSurf, rec2, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    'ElseIf Npc(MapResource)(MapResourceNum).Num).Big = 3 Then
        'Call DD_BackBuffer.BltFast(x, y, DD_BuildingSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
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

Sub ProcessPushBlock(ByVal x As Integer, ByVal y As Integer)
    ' Check if player is walking, and if so process moving the pushblock over
    If PushTile(x, y).Moving = MOVING_WALKING Then
        If PushTile(x, y).Pushed = YES Then
            Select Case PushTile(x, y).Dir
                Case DIR_UP
                    PushTile(x, y).YOffset = PushTile(x, y).YOffset - WALK_SPEED
                    ' Check if completed walking over to the next tile
                    If (PushTile(x, y).XOffset = 0) And (PushTile(x, y).YOffset = -32) Then
                        PushTile(x, y).Moving = 0
                        Map.Tile(x, y - 1).Mask = Map.Tile(x, y).Mask
                    End If
                Case DIR_DOWN
                    PushTile(x, y).YOffset = PushTile(x, y).YOffset + WALK_SPEED
                    ' Check if completed walking over to the next tile
                    If (PushTile(x, y).XOffset = 0) And (PushTile(x, y).YOffset = 32) Then
                        PushTile(x, y).Moving = 0
                        Map.Tile(x, y + 1).Mask = Map.Tile(x, y).Mask
                    End If
                Case DIR_LEFT
                    PushTile(x, y).XOffset = PushTile(x, y).XOffset - WALK_SPEED
                    ' Check if completed walking over to the next tile
                    If (PushTile(x, y).XOffset = -32) And (PushTile(x, y).YOffset = 0) Then
                        PushTile(x, y).Moving = 0
                        Map.Tile(x - 1, y).Mask = Map.Tile(x, y).Mask
                    End If
                Case DIR_RIGHT
                    PushTile(x, y).XOffset = PushTile(x, y).XOffset + WALK_SPEED
                    ' Check if completed walking over to the next tile
                    If (PushTile(x, y).XOffset = 32) And (PushTile(x, y).YOffset = 0) Then
                        PushTile(x, y).Moving = 0
                        Map.Tile(x + 1, y).Mask = Map.Tile(x, y).Mask
                    End If
            End Select
        Else
            Select Case PushTile(x, y).Dir
                Case DIR_UP
                    PushTile(x, y).YOffset = PushTile(x, y).YOffset + WALK_SPEED
                    ' Check if completed walking over to the next tile
                    If (PushTile(x, y).XOffset = 0) And (PushTile(x, y).YOffset = 0) Then
                        PushTile(x, y).Moving = 0
                        'Map.Tile(x, y - 1).Mask = Map.Tile(x, y).Mask
                    End If
                Case DIR_DOWN
                    PushTile(x, y).YOffset = PushTile(x, y).YOffset - WALK_SPEED
                    ' Check if completed walking over to the next tile
                    If (PushTile(x, y).XOffset = 0) And (PushTile(x, y).YOffset = 0) Then
                        PushTile(x, y).Moving = 0
                        'Map.Tile(x, y + 1).Mask = Map.Tile(x, y).Mask
                    End If
                Case DIR_LEFT
                    PushTile(x, y).XOffset = PushTile(x, y).XOffset + WALK_SPEED
                    ' Check if completed walking over to the next tile
                    If (PushTile(x, y).XOffset = 0) And (PushTile(x, y).YOffset = 0) Then
                        PushTile(x, y).Moving = 0
                        'Map.Tile(x - 1, y).Mask = Map.Tile(x, y).Mask
                    End If
                Case DIR_RIGHT
                    PushTile(x, y).XOffset = PushTile(x, y).XOffset - WALK_SPEED
                    ' Check if completed walking over to the next tile
                    If (PushTile(x, y).XOffset = 0) And (PushTile(x, y).YOffset = 0) Then
                        PushTile(x, y).Moving = 0
                        'Map.Tile(x + 1, y).Mask = Map.Tile(x, y).Mask
                    End If
            End Select
        End If
    End If

    ' Check if player is running, and if so process moving the pushblock over
    If PushTile(x, y).Moving = MOVING_RUNNING Then
        Select Case PushTile(x, y).Dir
            Case DIR_UP
                PushTile(x, y).YOffset = PushTile(x, y).YOffset - RUN_SPEED
                ' Check if completed walking over to the next tile
                If (PushTile(x, y).XOffset = 0) And (PushTile(x, y).YOffset = -32) Then
                    PushTile(x, y).Moving = 0
                    Map.Tile(x, y - 1).Mask = Map.Tile(x, y).Mask
                End If
            Case DIR_DOWN
                PushTile(x, y).YOffset = PushTile(x, y).YOffset + RUN_SPEED
                ' Check if completed walking over to the next tile
                If (PushTile(x, y).XOffset = 0) And (PushTile(x, y).YOffset = 32) Then
                    PushTile(x, y).Moving = 0
                    Map.Tile(x, y + 1).Mask = Map.Tile(x, y).Mask
                End If
            Case DIR_LEFT
                PushTile(x, y).XOffset = PushTile(x, y).XOffset - RUN_SPEED
                ' Check if completed walking over to the next tile
                If (PushTile(x, y).XOffset = -32) And (PushTile(x, y).YOffset = 0) Then
                    PushTile(x, y).Moving = 0
                    Map.Tile(x - 1, y).Mask = Map.Tile(x, y).Mask
                End If
            Case DIR_RIGHT
                PushTile(x, y).XOffset = PushTile(x, y).XOffset + RUN_SPEED
                ' Check if completed walking over to the next tile
                If (PushTile(x, y).XOffset = 32) And (PushTile(x, y).YOffset = 0) Then
                    PushTile(x, y).Moving = 0
                    Map.Tile(x + 1, y).Mask = Map.Tile(x, y).Mask
                End If
        End Select
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

Sub CheckInput2()
    If GettingMap = False Then
        'If GetKeyState(VK_RETURN) < 0 Then
            'Call CheckMapGetItem
        'End If
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
            'If KeyCode = vbKeyReturn Then
                'Call CheckMapGetItem
            'End If
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

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
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
            Call AddText("Available Commands: /help, /who", HelpColor) ' /info, /who, /fps, /inv, /stats, /train, /trade, /party, /join, /leave", HelpColor)
            MyText = ""
            Exit Sub
        End If
        
        '' Verification User
        'If LCase(Mid(MyText, 1, 5)) = "/info" Then
            'ChatText = Mid(MyText, 6, Len(MyText) - 5)
            'Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            'MyText = ""
            'Exit Sub
        'End If
        
        ' Whos Online
        If LCase(Mid(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = ""
            Exit Sub
        End If
                        
        '' Checking fps
        'If LCase(Mid(MyText, 1, 4)) = "/fps" Then
            'Call AddText("FPS: " & GameFPS, Pink)
            'MyText = ""
            'Exit Sub
        'End If
                
        '' Show inventory
        'If LCase(Mid(MyText, 1, 4)) = "/inv" Then
            'Call UpdateInventory
            'frmMirage.picInv.Visible = True
            'MyText = ""
            'Exit Sub
        'End If
        
        '' Request stats
        'If LCase(Mid(MyText, 1, 6)) = "/stats" Then
            'Call SendData("getstats" & SEP_CHAR & END_CHAR)
            'MyText = ""
            'Exit Sub
        'End If
    
        '' Show training
        'If LCase(Mid(MyText, 1, 6)) = "/train" Then
            'frmTraining.Show vbModal
            'MyText = ""
            'Exit Sub
        'End If

        '' Request trade
        'If LCase(Mid(MyText, 1, 6)) = "/trade" Then
            'Call SendData("trade" & SEP_CHAR & END_CHAR)
            'MyText = ""
            'Exit Sub
        'End If
        
         '' Party request
7        'If LCase(Mid(MyText, 1, 6)) = "/party" Then
            '' Make sure they are actually sending something
            'If Len(MyText) > 7 Then
                'ChatText = Mid(MyText, 8, Len(MyText) - 7)
                'Call SendPartyRequest(ChatText)
            'Else
                'Call AddText("Usage: /party playernamehere", AlertColor)
            'End If
            'MyText = ""
            'Exit Sub
        'End If
        
        '' Join party
        'If LCase(Mid(MyText, 1, 5)) = "/join" Then
            'Call SendJoinParty
            'MyText = ""
            'Exit Sub
        'End If
        
        '' Leave party
        'If LCase(Mid(MyText, 1, 6)) = "/leave" Then
            'Call SendLeaveParty
            'MyText = ""
            'Exit Sub
        'End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' Admin Help
            If LCase(Mid(MyText, 1, 6)) = "/admin" Then
                Call AddText("Social Commands:", HelpColor)
                Call AddText("""msghere = Global Admin Message", HelpColor)
                Call AddText("=msghere = Private Admin Message", HelpColor)
                Call AddText("Available Commands: /admin, /mapeditor, /edititem, /respawn, /editnpc, /editshop, /editspell, /editskill, /editquest, /editmenu", HelpColor)
                MyText = ""
                Exit Sub
            End If
            
            '' Kicking a player
            'If LCase(Mid(MyText, 1, 5)) = "/kick" Then
                'If Len(MyText) > 6 Then
                    'MyText = Mid(MyText, 7, Len(MyText) - 6)
                    'Call SendKick(MyText)
                'End If
                'MyText = ""
                'Exit Sub
            'End If
        
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
            '' Location
            'If LCase(Mid(MyText, 1, 4)) = "/loc" Then
                'Call SendRequestLocation
                'MyText = ""
                'Exit Sub
            'End If
            
            ' Map Editor
            If LCase(Mid(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                MyText = ""
                Exit Sub
            End If
            
            '' Warping to a player
            'If LCase(Mid(MyText, 1, 9)) = "/warpmeto" Then
                  'If Len(MyText) > 10 Then
                    'MyText = Mid(MyText, 10, Len(MyText) - 9)
                    'Call WarpMeTo(MyText)
                'End If
                'MyText = ""
                'Exit Sub
            'End If
                        
            '' Warping a player to you
            'If LCase(Mid(MyText, 1, 9)) = "/warptome" Then
                'If Len(MyText) > 10 Then
                    'MyText = Mid(MyText, 10, Len(MyText) - 9)
                    'Call WarpToMe(MyText)
                'End If
                'MyText = ""
                'Exit Sub
            'End If
                        
            '' Warping to a map
            'If LCase(Mid(MyText, 1, 7)) = "/warpto" Then
                'If Len(MyText) > 8 Then
                    'MyText = Mid(MyText, 8, Len(MyText) - 7)
                    'n = Val(MyText)
                
                    '' Check to make sure its a valid map #
                    'If n > 0 And n <= MAX_MAPS Then
                        'Call WarpTo(n)
                    'Else
                        'Call AddText("Invalid map number.", Red)
                    'End If
                'End If
                'MyText = ""
                'Exit Sub
            'End If
            
            '' Setting sprite
            'If LCase(Mid(MyText, 1, 10)) = "/setsprite" Then
                'If Len(MyText) > 11 Then
                    '' Get sprite #
                    'MyText = Mid(MyText, 12, Len(MyText) - 11)
                
                    'Call SendSetSprite(Val(MyText))
                'End If
                'MyText = ""
                'Exit Sub
            'End If
            
            '' Map report
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
        
            '' MOTD change
            'If Mid(MyText, 1, 5) = "/motd" Then
                'If Len(MyText) > 6 Then
                    'MyText = Mid(MyText, 7, Len(MyText) - 6)
                    'If Trim(MyText) <> "" Then
                        'Call SendMOTDChange(MyText)
                    'End If
                'End If
                'MyText = ""
                'Exit Sub
            'End If
            
            '' Check the ban list
            'If Mid(MyText, 1, 3) = "/banlist" Then
                'Call SendBanList
                'MyText = ""
                'Exit Sub
            'End If
            
            '' Banning a player
            'If LCase(Mid(MyText, 1, 4)) = "/ban" Then
                'If Len(MyText) > 5 Then
                    'MyText = Mid(MyText, 6, Len(MyText) - 5)
                    'Call SendBan(MyText)
                    'MyText = ""
                'End If
                'Exit Sub
            'End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
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
            If Mid(MyText, 1, 10) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = ""
                Exit Sub
            End If
            
            ' Editing skill request
            If Mid(MyText, 1, 10) = "/editskill" Then
                Call SendRequestEditSkill
                MyText = ""
                Exit Sub
            End If
            
            ' Editing quest request
            If Mid(MyText, 1, 10) = "/editquest" Then
                Call SendRequestEditQuest
                MyText = ""
                Exit Sub
            End If
            
            ' Editing Menu request
            If Mid(MyText, 1, 9) = "/editmenu" Then
                Call SendRequestEditGUI
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
            
            '' Ban destroy
            'If LCase(Mid(MyText, 1, 15)) = "/destroybanlist" Then
                'Call SendBanDestroy
                'MyText = ""
                'Exit Sub
            'End If
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
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_KEY Then
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
            
            ' If there's a pushblock there, see if we can push it
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_PUSHBLOCK Then
                If PushTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Pushed = NO Then
                    If (Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data1 <> DIR_UP + 1) And (Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data2 <> DIR_UP + 1) And (Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data3 <> DIR_UP + 1) Then
                        CanMove = False
                    
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            End If
            
            ' Check to see if a player is already on that tile
            For i = 1 To HighIndex
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
            
            ' Check to see if a resource is already on that tile
            For i = 1 To MAX_MAP_RESOURCES
                If MapResource(i).Num > 0 Then
                    If (MapResource(i).x = GetPlayerX(MyIndex)) And (MapResource(i).y = GetPlayerY(MyIndex) - 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
                   
            ' Check if the tile will let us walk onto it
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).WalkUp = 1 Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
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
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_KEY Then
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
            
            ' If there's a pushblock there, see if we can push it
            If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_PUSHBLOCK Then
                If PushTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Pushed = NO Then
                    If (Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data1 <> DIR_DOWN + 1) And (Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data2 <> DIR_DOWN + 1) And (Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data3 <> DIR_DOWN + 1) Then
                        CanMove = False
                    
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            End If
            
            ' Check to see if a player is already on that tile
            For i = 1 To HighIndex
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
            
            ' Check to see if a resource is already on that tile
            For i = 1 To MAX_MAP_RESOURCES
                If MapResource(i).Num > 0 Then
                    If (MapResource(i).x = GetPlayerX(MyIndex)) And (MapResource(i).y = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            ' Check if the tile will let us walk onto it
             If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).WalkDown = 1 Then
                 CanMove = False
                
                 ' Set the new direction if they weren't facing that direction
                 If d <> DIR_DOWN Then
                     Call SendPlayerDir
                 End If
                 Exit Function
             End If
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
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    'Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Then
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
            
            ' If there's a pushblock there, see if we can push it
            If Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_PUSHBLOCK Then
                If PushTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Pushed = NO Then
                    If (Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data1 <> DIR_LEFT + 1) And (Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data2 <> DIR_LEFT + 1) And (Map.Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data3 <> DIR_LEFT + 1) Then
                        CanMove = False
                    
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            End If
            
            ' Check to see if a player is already on that tile
            For i = 1 To HighIndex
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
            
            ' Check to see if a resource is already on that tile
            For i = 1 To MAX_MAP_RESOURCES
                If MapResource(i).Num > 0 Then
                    If (MapResource(i).x = GetPlayerX(MyIndex) - 1) And (MapResource(i).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            ' Check if the tile will let us walk onto it
             If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).WalkLeft = 1 Then
                 CanMove = False
                
                 ' Set the new direction if they weren't facing that direction
                 If d <> DIR_LEFT Then
                     Call SendPlayerDir
                 End If
                 Exit Function
             End If
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
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Then
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
            
            ' If there's a pushblock there, see if we can push it
            If Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_PUSHBLOCK Then
                If PushTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Pushed = NO Then
                    If (Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data1 <> DIR_RIGHT + 1) And (Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data2 <> DIR_RIGHT + 1) And (Map.Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data3 <> DIR_RIGHT + 1) Then
                        CanMove = False
                    
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            End If
            
            ' Check to see if a player is already on that tile
            For i = 1 To HighIndex
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
            
            ' Check to see if a resource is already on that tile
            For i = 1 To MAX_MAP_RESOURCES
                If MapResource(i).Num > 0 Then
                    If (MapResource(i).x = GetPlayerX(MyIndex) + 1) And (MapResource(i).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            ' Check if the tile will let us walk onto it
             If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).WalkRight = 1 Then
                 CanMove = False
                
                 ' Set the new direction if they weren't facing that direction
                 If d <> DIR_RIGHT Then
                     Call SendPlayerDir
                 End If
                 Exit Function
             End If
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

Public Sub EditorInit()
    
    SaveMap = Map
    InEditor = True
    RSpawnNum = 20
    frmCClient.picRightClickMenu.Visible = False
    frmCClient.picMapEditor.Visible = True
    With frmCClient.picBackSelect
        .Width = 14 * PIC_X
        .Height = 255 * PIC_Y
        .Picture = LoadPicture(App.Path + GFX_PATH + "Tiles" + GFX_EXT)
    End With
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long ' Used to record where on the tile was clicked
Dim CenterTolerence As RECT ' Controls how easy it is to press center

    If InEditor Then
        x1 = Int(x / PIC_X)
        y1 = Int(y / PIC_Y)
        ' Check if we need to change directional block or tile attribs
        If frmCClient.optDirectionView.Value = True Then
            With CenterTolerence
                .Top = 9
                .Bottom = 22
                .Left = 9
                .Right = 22
                
                If Button = 1 And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                    x2 = x - (x1 * PIC_X)
                    y2 = y - (y1 * PIC_Y)
                    
                    If Shift = 1 Then
                        ' Using CenterTolerence as a grid guide, check which part was clicked (Start from Bottom right)
                        If x2 > .Right Then ' Right side
                            If y2 > .Bottom Then ' Right Bottom
                                ' Now check which side of that small section was clicked
                                If x2 > y2 Then
                                    Map.Tile(x1, y1).WalkRight = 0
                                  
                                Else
                                    Map.Tile(x1, y1).WalkDown = 0
                                End If
                             
                            ElseIf y2 > .Left Then ' Right Middle
                            Map.Tile(x1, y1).WalkRight = 0
                             
                            Else ' Right Top
                                ' Check which side was clicked, remember, minus from x because
                                ' its not starting from 0
                                If x2 - .Right > y2 Then
                                    Map.Tile(x1, y1).WalkUp = 0
                                  
                                Else
                                    Map.Tile(x1, y1).WalkRight = 0
                                End If
                             
                            End If
                         
                        ElseIf x2 > .Left Then ' Middle side
                            If y2 > .Bottom Then 'Bottom
                                Map.Tile(x1, y1).WalkDown = 0
                             
                            ElseIf y2 > .Left Then ' Middle
                                Map.Tile(x1, y1).WalkUp = 0
                                Map.Tile(x1, y1).WalkDown = 0
                                Map.Tile(x1, y1).WalkLeft = 0
                                Map.Tile(x1, y1).WalkRight = 0
                             
                            Else ' Top
                                Map.Tile(x1, y1).WalkUp = 0
                            End If
                         
                        Else
                            If y2 > .Bottom Then 'Left Bottom
                                If x2 > y2 - .Bottom Then
                                    Map.Tile(x1, y1).WalkDown = 0
                                Else
                                    Map.Tile(x1, y1).WalkLeft = 0
                                End If
                             
                            ElseIf y2 > .Left Then ' Left Middle
                                Map.Tile(x1, y1).WalkLeft = 0
                             
                            Else ' Left Top
                                If x2 > y2 Then
                                    Map.Tile(x1, y1).WalkUp = 0
                                Else
                                    Map.Tile(x1, y1).WalkLeft = 0
                                End If
                            End If
                         
                        End If
                    Else
                        ' Using CenterTolerence as a grid guide, check which part was clicked (Start from Bottom right)
                        If x2 > .Right Then ' Right side
                            If y2 > .Bottom Then ' Right Bottom
                                ' Now check which side of that small section was clicked
                                If x2 > y2 Then
                                    Map.Tile(x1, y1).WalkRight = 1
                                  
                                Else
                                    Map.Tile(x1, y1).WalkDown = 1
                                End If
                             
                            ElseIf y2 > .Left Then ' Right Middle
                                Map.Tile(x1, y1).WalkRight = 1
                             
                            Else ' Right Top
                            ' Check which side was clicked, remember, minus from x because
                            ' its not starting from 0
                                If x2 - .Right > y2 Then
                                Map.Tile(x1, y1).WalkUp = 1
                                     
                                Else
                                    Map.Tile(x1, y1).WalkRight = 1
                                End If
                             
                            End If
                         
                        ElseIf x2 > .Left Then ' Middle side
                            If y2 > .Bottom Then 'Bottom
                                Map.Tile(x1, y1).WalkDown = 1
                             
                            ElseIf y2 > .Left Then ' Middle
                                Map.Tile(x1, y1).WalkUp = 1
                                Map.Tile(x1, y1).WalkDown = 1
                                Map.Tile(x1, y1).WalkLeft = 1
                                Map.Tile(x1, y1).WalkRight = 1
                             
                            Else ' Top
                                Map.Tile(x1, y1).WalkUp = 1
                            End If
                         
                        Else
                            If y2 > .Bottom Then 'Left Bottom
                                If x2 > y2 - .Bottom Then
                                    Map.Tile(x1, y1).WalkDown = 1
                                Else
                                    Map.Tile(x1, y1).WalkLeft = 1
                                End If
                             
                            ElseIf y2 > .Left Then ' Left Middle
                                Map.Tile(x1, y1).WalkLeft = 1
                             
                            Else ' Left Top
                                If x2 > y2 Then
                                    Map.Tile(x1, y1).WalkUp = 1
                                Else
                                    Map.Tile(x1, y1).WalkLeft = 1
                                End If
                            End If
                         
                        End If
                    End If
                End If
            End With
        Else
            If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
                If frmCClient.optLayers.Value = True Then
                    If Shift = 1 Then
                        With Map.Tile(x1, y1)
                            If frmCClient.optGround.Value = True Then .Ground = 0
                            If frmCClient.optMask.Value = True Then .Mask = 0
                            If frmCClient.optMask2.Value = True Then .Mask2 = 0
                            If frmCClient.optAnim.Value = True Then .Anim = 0
                            If frmCClient.optFringe.Value = True Then .Fringe = 0
                            If frmCClient.optFringe2.Value = True Then .Fringe2 = 0
                            If frmCClient.optFAnim.Value = True Then .FAnim = 0
                            If frmCClient.optLight.Value = True Then .Light = 0
                        End With
                    Else
                        With Map.Tile(x1, y1)
                            If frmCClient.optGround.Value = True Then .Ground = EditorTileY * 14 + EditorTileX
                            If frmCClient.optMask.Value = True Then .Mask = EditorTileY * 14 + EditorTileX
                            If frmCClient.optMask2.Value = True Then .Mask2 = EditorTileY * 14 + EditorTileX
                            If frmCClient.optAnim.Value = True Then .Anim = EditorTileY * 14 + EditorTileX
                            If frmCClient.optFringe.Value = True Then .Fringe = EditorTileY * 14 + EditorTileX
                            If frmCClient.optFringe2.Value = True Then .Fringe2 = EditorTileY * 14 + EditorTileX
                            If frmCClient.optFAnim.Value = True Then .FAnim = EditorTileY * 14 + EditorTileX
                            If frmCClient.optLight.Value = True Then .Light = EditorTileY * 14 + EditorTileX
                        End With
                    End If
                ElseIf frmCClient.optAttribs.Value = True Then
                    If Shift = 1 Then
                        With Map.Tile(x1, y1)
                            .Type = 0
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                        End With
                    Else
                        With Map.Tile(x1, y1)
                            If frmCClient.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                            If frmCClient.optWarp.Value = True Then
                                .Type = TILE_TYPE_WARP
                                .Data1 = EditorWarpMap
                                .Data2 = EditorWarpX
                                .Data3 = EditorWarpY
                            End If
                            If frmCClient.optItem.Value = True Then
                                .Type = TILE_TYPE_ITEM
                                .Data1 = ItemEditorNum
                                .Data2 = ItemEditorValue
                                .Data3 = 0
                            End If
                            If frmCClient.optNpcAvoid.Value = True Then
                                .Type = TILE_TYPE_NPCAVOID
                                .Data1 = 0
                                .Data2 = 0
                                .Data3 = 0
                            End If
                            If frmCClient.optKey.Value = True Then
                                .Type = TILE_TYPE_KEY
                                .Data1 = KeyEditorNum
                                .Data2 = KeyEditorTake
                                .Data3 = 0
                            End If
                            If frmCClient.optKeyOpen.Value = True Then
                                .Type = TILE_TYPE_KEYOPEN
                                .Data1 = KeyOpenEditorX
                                .Data2 = KeyOpenEditorY
                                .Data3 = 0
                            End If
                            
                            If frmCClient.optPushBlock.Value = True Then
                                .Type = TILE_TYPE_PUSHBLOCK
                                .Data1 = PushDir1
                                .Data2 = PushDir2
                                .Data3 = PushDir3
                                Map.Tile(x1, y1 - 1).Type = TILE_TYPE_NPCAVOID
                                Map.Tile(x1, y1 + 1).Type = TILE_TYPE_NPCAVOID
                                Map.Tile(x1 - 1, y1).Type = TILE_TYPE_NPCAVOID
                                Map.Tile(x1 + 1, y1).Type = TILE_TYPE_NPCAVOID
                            End If
                        End With
                    End If
                Else
                    If Shift = 1 Then
                        With Map.Tile(x1, y1)
                            .Build = 0
                        End With
                    Else
                        With Map.Tile(x1, y1)
                            .Build = 1
                        End With
                    End If
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
    Call BitBlt(frmCClient.picSelect.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picBackSelect.hdc, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub EditorTileScroll()
    frmCClient.picBackSelect.Top = (frmCClient.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    Map = SaveMap
    InEditor = False
    frmCClient.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, x As Long, y As Long

    ' Ground layer
    If frmCClient.optGround.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map.Tile(x, y).Ground = 0
                Next x
            Next y
        End If
    End If

    ' Mask layer
    If frmCClient.optMask.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map.Tile(x, y).Mask = 0
                Next x
            Next y
        End If
    End If

    ' Mask 2 layer
    If frmCClient.optMask2.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, GAME_NAME)
       
        If YesNo = vbYes Then
             For y = 0 To MAX_MAPY
                 For x = 0 To MAX_MAPX
                     Map.Tile(x, y).Mask2 = 0
                 Next x
             Next y
        End If
    End If

    ' Animation layer
    If frmCClient.optAnim.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    Map.Tile(x, y).Anim = 0
                Next x
            Next y
        End If
    End If

    ' Fringe layer
    If frmCClient.optFringe.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
       
        If YesNo = vbYes Then
             For y = 0 To MAX_MAPY
                 For x = 0 To MAX_MAPX
                     Map.Tile(x, y).Fringe = 0
                 Next x
             Next y
        End If
    End If

    ' Fringe 2 layer
    If frmCClient.optFringe2.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, GAME_NAME)
       
        If YesNo = vbYes Then
             For y = 0 To MAX_MAPY
                 For x = 0 To MAX_MAPX
                     Map.Tile(x, y).Fringe2 = 0
                 Next x
             Next y
        End If
    End If

    ' Fringe Animation layer
    If frmCClient.optFAnim.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)
       
        If YesNo = vbYes Then
             For y = 0 To MAX_MAPY
                 For x = 0 To MAX_MAPX
                     Map.Tile(x, y).FAnim = 0
                 Next x
             Next y
        End If
    End If
    
    ' Light layer
    If frmCClient.optLight.Value = True Then
        YesNo = MsgBox("Are you sure you wish to clear the Light layer?", vbYesNo, GAME_NAME)
       
        If YesNo = vbYes Then
             For y = 0 To MAX_MAPY
                 For x = 0 To MAX_MAPX
                     Map.Tile(x, y).Light = 0
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
                Map.Tile(x, y).Type = 0
            Next x
        Next y
    End If
End Sub

Public Sub ItemEditorInit()
    frmItemEditor.picItemsBack.Picture = LoadPicture(App.Path & GFX_PATH & "Items" & GFX_EXT)
    frmItemEditor.picArrowsBack.Picture = LoadPicture(App.Path & GFX_PATH & "Arrows" & GFX_EXT)
    frmItemEditor.tmrItemPic.Enabled = True
    
    frmItemEditor.txtItemName.Text = Trim(Item(EditorIndex).Name)
    frmItemEditor.scrlItemPic.Max = MAX_ITEMS
    frmItemEditor.scrlItemPic.Value = Item(EditorIndex).Pic
    frmItemEditor.cmbItemType.ListIndex = Item(EditorIndex).Type
    
    If (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_TOOL) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.cmbItemSubType.ListIndex = Item(EditorIndex).Data3
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If Not (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_TOOL) Then
        If (frmItemEditor.cmbItemType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbItemType.ListIndex <= ITEM_TYPE_SHIELD) Then
            frmItemEditor.fraEquipment.Visible = True
            frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
            frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
            frmItemEditor.cmbItemSubType.ListIndex = Item(EditorIndex).Data3
        Else
            frmItemEditor.fraEquipment.Visible = False
        End If
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex >= ITEM_TYPE_AMULET) And (frmItemEditor.cmbItemType.ListIndex <= ITEM_TYPE_RING) Then
        frmItemEditor.fraCharm.Visible = True
        frmItemEditor.cmbCharmType.ListIndex = Item(EditorIndex).Data1
        frmItemEditor.scrlCharmMod.Value = Item(EditorIndex).Data2
    Else
        frmItemEditor.fraCharm.Visible = False
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbItemType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Enabled = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraSpell.Visible = False
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_SKILL) Then
        frmItemEditor.fraSkill.Visible = True
        frmItemEditor.scrlSkill.Enabled = True
        frmItemEditor.scrlSkill.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraSkill.Visible = False
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_ARROW) Then
        frmItemEditor.fraArrows.Visible = True
        frmItemEditor.scrlArrowRange.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlArrowQuantity.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlArrowAnim.Value = Item(EditorIndex).Data3
    Else
        frmItemEditor.fraArrows.Visible = False
    End If
    
    frmItemEditor.Visible = True
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).Name = frmItemEditor.txtItemName.Text
    Item(EditorIndex).Pic = frmItemEditor.scrlItemPic.Value
    Item(EditorIndex).Type = frmItemEditor.cmbItemType.ListIndex

    If (frmItemEditor.cmbItemType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbItemType.ListIndex <= ITEM_TYPE_SHIELD) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = frmItemEditor.cmbItemSubType.ListIndex
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_TOOL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = frmItemEditor.cmbItemSubType.ListIndex
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex >= ITEM_TYPE_AMULET) And (frmItemEditor.cmbItemType.ListIndex <= ITEM_TYPE_RING) Then
        Item(EditorIndex).Data1 = frmItemEditor.cmbCharmType.ListIndex
        Item(EditorIndex).Data2 = frmItemEditor.scrlCharmMod.Value
        Item(EditorIndex).Data3 = 0
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbItemType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_SKILL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSkill.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    If (frmItemEditor.cmbItemType.ListIndex = ITEM_TYPE_ARROW) Then
        Item(EditorIndex).Data2 = frmItemEditor.scrlArrowRange.Value
        Item(EditorIndex).Data1 = frmItemEditor.scrlArrowQuantity.Value
        Item(EditorIndex).Data3 = frmItemEditor.scrlArrowAnim.Value
    End If
    
    If Trim(Item(EditorIndex).Name) = "" Then
        Item(EditorIndex).Type = 0
        Item(EditorIndex).Data1 = 0
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
    End If
    
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    frmItemEditor.tmrItemPic.Enabled = False
    frmItemEditor.Visible = False
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    frmItemEditor.tmrItemPic.Enabled = False
    frmItemEditor.Visible = False
End Sub

Public Sub ItemEditorBltItem()
    Call BitBlt(frmItemEditor.picItemPic.hdc, 0, 0, PIC_X, PIC_Y, frmItemEditor.picItemsBack.hdc, 0, frmItemEditor.scrlItemPic.Value * PIC_Y, SRCCOPY)
    Call BitBlt(frmItemEditor.picArrowAnim.hdc, 0, 0, PIC_X, PIC_Y, frmItemEditor.picArrowsBack.hdc, PIC_X, frmItemEditor.scrlArrowAnim.Value * PIC_Y, SRCCOPY)
End Sub

Public Sub NpcEditorInit()
Dim i As Long
Dim n As Long
    
    frmEditNpc.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    frmEditNpc.tmrNpcSprite.Enabled = True
    
    frmEditNpc.txtNpcName.Text = Trim(Npc(EditorIndex).Name)
    frmEditNpc.scrlNpcSprite.Max = MAX_NPCS
    frmEditNpc.scrlNpcSprite.Value = Npc(EditorIndex).Sprite
    frmEditNpc.txtNpcSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmEditNpc.cmbNpcBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmEditNpc.scrlNpcRange.Value = Npc(EditorIndex).Range
    frmEditNpc.scrlNpcSTR.Value = Npc(EditorIndex).STR
    frmEditNpc.scrlNpcDEF.Value = Npc(EditorIndex).DEF
    frmEditNpc.scrlNpcSPEED.Value = Npc(EditorIndex).SPEED
    frmEditNpc.scrlNpcMAGI.Value = Npc(EditorIndex).MAGI
    If Npc(EditorIndex).Big = 1 Then
        frmEditNpc.chkBigNpc.Value = Checked
        frmEditNpc.chkTreeNpc.Value = Unchecked
        frmEditNpc.chkBuildingNpc.Value = Unchecked
    ElseIf Npc(EditorIndex).Big = 2 Then
        frmEditNpc.chkBigNpc.Value = Unchecked
        frmEditNpc.chkTreeNpc = Checked
        frmEditNpc.chkBuildingNpc.Value = Unchecked
    ElseIf Npc(EditorIndex).Big = 3 Then
        frmEditNpc.chkBigNpc.Value = Unchecked
        frmEditNpc.chkTreeNpc = Unchecked
        frmEditNpc.chkBuildingNpc.Value = Checked
    Else
        frmEditNpc.chkBigNpc.Value = Unchecked
        frmEditNpc.chkTreeNpc = Unchecked
        frmEditNpc.chkBuildingNpc.Value = Unchecked
    End If
    frmEditNpc.scrlNpcStartHP.Value = Npc(EditorIndex).MaxHp
    frmEditNpc.cmbExpType.Clear
    frmEditNpc.cmbExpType.AddItem "Standard", 0
    For i = 1 To MAX_SKILLS
        frmEditNpc.cmbExpType.AddItem i & ": " & Trim(Skill(i).Name)
    Next i
    frmEditNpc.cmbExpType.ListIndex = Npc(EditorIndex).ExpType
    frmEditNpc.scrlNpcExpGiven.Value = Npc(EditorIndex).EXP
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
        frmEditNpc.chkNpcRespawn.Value = Unchecked
        frmEditNpc.chkHitOnlyWith.Value = Unchecked
        frmEditNpc.chkLinkWithShop.Value = Checked
    ElseIf Npc(EditorIndex).Behavior = NPC_BEHAVIOR_FRIENDLY Then
        frmEditNpc.chkNpcRespawn.Value = Unchecked
        frmEditNpc.chkHitOnlyWith.Value = Unchecked
        frmEditNpc.chkLinkWithShop.Value = Unchecked
    Else
        frmEditNpc.chkNpcRespawn.Value = Checked
        frmEditNpc.chkHitOnlyWith.Value = Checked
        frmEditNpc.chkLinkWithShop.Value = Unchecked
    End If
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SHOPKEEPER Or Npc(EditorIndex).Behavior = NPC_BEHAVIOR_FRIENDLY Then
        frmEditNpc.cmbNpcRespawn.Clear
        frmEditNpc.cmbNpcRespawn.AddItem "None", 0
        frmEditNpc.cmbNpcRespawn.ListIndex = 0
    Else
        frmEditNpc.cmbNpcRespawn.Clear
        frmEditNpc.cmbNpcRespawn.AddItem "None", 0
        For i = 1 To MAX_NPCS
            frmEditNpc.cmbNpcRespawn.AddItem i & ": " & Trim(Npc(i).Name)
        Next i
        frmEditNpc.cmbNpcRespawn.ListIndex = Npc(EditorIndex).Respawn
    End If
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SHOPKEEPER Or Npc(EditorIndex).Behavior = NPC_BEHAVIOR_FRIENDLY Then
        frmEditNpc.cmbHitOnlyWith.Clear
        frmEditNpc.cmbHitOnlyWith.AddItem "None", 0
        frmEditNpc.cmbHitOnlyWith.ListIndex = 0
    Else
        frmEditNpc.cmbHitOnlyWith.Clear
        frmEditNpc.cmbHitOnlyWith.AddItem "None", 0
        If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_RESOURCE Then
            For i = 1 To MAX_ITEMS
                If Item(i).Type = ITEM_TYPE_TOOL Then
                    frmEditNpc.cmbHitOnlyWith.AddItem i & ": " & Trim(Item(i).Name)
                Else
                    frmEditNpc.cmbHitOnlyWith.AddItem i & ": " & "Not Tool Type"
                End If
            Next i
        Else
            For i = 1 To MAX_ITEMS
                If Item(i).Type = ITEM_TYPE_WEAPON Then
                    frmEditNpc.cmbHitOnlyWith.AddItem i & ": " & Trim(Item(i).Name)
                Else
                    frmEditNpc.cmbHitOnlyWith.AddItem i & ": " & "Not Weapon Type"
                End If
            Next i
        End If
        frmEditNpc.cmbHitOnlyWith.ListIndex = Npc(EditorIndex).HitOnlyWith
    End If
    If Npc(EditorIndex).Behavior = NPC_BEHAVIOR_SHOPKEEPER Then
        frmEditNpc.cmbShopLink.Clear
        frmEditNpc.cmbShopLink.AddItem "None", 0
        For i = 1 To MAX_SHOPS
            frmEditNpc.cmbShopLink.AddItem i & ": " & Trim(Shop(i).Name)
        Next i
        frmEditNpc.cmbShopLink.ListIndex = Npc(EditorIndex).ShopLink
    Else
        frmEditNpc.cmbShopLink.Clear
        frmEditNpc.cmbShopLink.AddItem "None", 0
        frmEditNpc.cmbShopLink.ListIndex = 0
    End If
    frmEditNpc.scrlNpcQuest.Value = 1
    frmEditNpc.scrlNpcQuestNum.Value = Npc(EditorIndex).QuestNPC(1)
    If Npc(EditorIndex).QuestNPC(1) > 0 Then
        frmEditNpc.lblNpcQuestName.Caption = Trim(Quest(Npc(EditorIndex).QuestNPC(1)).Name)
    Else
        frmEditNpc.lblNpcQuestName.Caption = "No Quest"
    End If
    frmEditNpc.lblNpcQuestNum.Caption = STR(frmEditNpc.scrlNpcQuestNum.Value)
    frmEditNpc.scrlDropItem.Max = MAX_NPC_DROPS
    frmEditNpc.scrlDropNum.Max = MAX_ITEMS
    frmEditNpc.txtDropChance.Text = STR(Npc(EditorIndex).ItemNPC(1).Chance)
    frmEditNpc.scrlDropNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmEditNpc.scrlDropValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
    
    frmEditNpc.Show vbModal
End Sub

Public Sub NpcEditorOk()
Dim i As Long

    Npc(EditorIndex).Name = frmEditNpc.txtNpcName.Text
    Npc(EditorIndex).Sprite = frmEditNpc.scrlNpcSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmEditNpc.txtNpcSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmEditNpc.cmbNpcBehavior.ListIndex
    Npc(EditorIndex).Range = frmEditNpc.scrlNpcRange.Value
    Npc(EditorIndex).STR = frmEditNpc.scrlNpcSTR.Value
    Npc(EditorIndex).DEF = frmEditNpc.scrlNpcDEF.Value
    Npc(EditorIndex).SPEED = frmEditNpc.scrlNpcSPEED.Value
    Npc(EditorIndex).MAGI = frmEditNpc.scrlNpcMAGI.Value
    If frmEditNpc.chkBigNpc.Value = Checked Then
        Npc(EditorIndex).Big = 1
    ElseIf frmEditNpc.chkTreeNpc.Value = Checked Then
        Npc(EditorIndex).Big = 2
    ElseIf frmEditNpc.chkBuildingNpc.Value = Checked Then
        Npc(EditorIndex).Big = 3
    Else
        Npc(EditorIndex).Big = 0
    End If
    Npc(EditorIndex).MaxHp = frmEditNpc.scrlNpcStartHP.Value
    Npc(EditorIndex).ExpType = frmEditNpc.cmbExpType.ListIndex
    Npc(EditorIndex).EXP = frmEditNpc.scrlNpcExpGiven.Value
    If frmEditNpc.cmbNpcRespawn.ListIndex > 0 Then
        Npc(EditorIndex).Respawn = frmEditNpc.cmbNpcRespawn.ListIndex
    Else
        Npc(EditorIndex).Respawn = 0
    End If
    If frmEditNpc.cmbHitOnlyWith.ListIndex > 0 Then
        Npc(EditorIndex).HitOnlyWith = frmEditNpc.cmbHitOnlyWith.ListIndex
    Else
        Npc(EditorIndex).HitOnlyWith = 0
    End If
    If frmEditNpc.cmbShopLink.ListIndex > 0 Then
        Npc(EditorIndex).ShopLink = frmEditNpc.cmbShopLink.ListIndex
    Else
        Npc(EditorIndex).ShopLink = 0
    End If
    
    Call SendSaveNpc(EditorIndex)
    InNpcEditor = False
    frmEditNpc.tmrNpcSprite.Enabled = False
    Unload frmEditNpc
End Sub
Public Sub NpcEditorCancel()
    InNpcEditor = False
    frmEditNpc.tmrNpcSprite.Enabled = False
    Unload frmEditNpc
End Sub

Public Sub NpcEditorBltSprite()
    If frmEditNpc.chkBigNpc.Value = Checked Then
        Call BitBlt(frmEditNpc.picSprite.hdc, 0, 0, 64, 64, frmEditNpc.picSprites.hdc, 3 * 64, frmEditNpc.scrlNpcSprite.Value * 64, SRCCOPY)
    ElseIf frmEditNpc.chkTreeNpc.Value = Checked Then
        Call BitBlt(frmEditNpc.picSprite.hdc, 0, 0, 96, 128, frmEditNpc.picSprites.hdc, 0, frmEditNpc.scrlNpcSprite.Value * 128, SRCCOPY)
    Else
        Call BitBlt(frmEditNpc.picSprite.hdc, 0, 0, PIC_X, PIC_Y * 2, frmEditNpc.picSprites.hdc, 3 * PIC_X, ((frmEditNpc.scrlNpcSprite.Value * 2) * PIC_Y), SRCCOPY)
    End If
End Sub

Public Sub ShopEditorInit()
Dim i As Long

    frmEditShop.txtShopName.Text = Trim(Shop(EditorIndex).Name)
    frmEditShop.chkShopFixesItems.Value = Shop(EditorIndex).FixesItems
    
    frmEditShop.scrlGiveItem.Max = MAX_GIVE_ITEMS
    frmEditShop.scrlGetItem.Max = MAX_GET_ITEMS
    
    frmEditShop.cmbShopItemGive.Clear
    frmEditShop.cmbShopItemGive.AddItem "None"
    frmEditShop.cmbShopItemGet.Clear
    frmEditShop.cmbShopItemGet.AddItem "None"
    For i = 1 To MAX_ITEMS
        frmEditShop.cmbShopItemGive.AddItem i & ": " & Trim(Item(i).Name)
        frmEditShop.cmbShopItemGet.AddItem i & ": " & Trim(Item(i).Name)
    Next i
    'frmEditShop.cmbShopItemGive.ListIndex = 0
    'frmEditShop.cmbShopItemGet.ListIndex = 0
    
    frmEditShop.lstShopTradeItem.ListIndex = 0
    Call UpdateShopTrade
    
    frmEditShop.Show vbModal
End Sub

'Public Sub UpdateShopTrade()
'Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long
    
    'frmEditShop.lstShopTradeItem.Clear
    'For i = 1 To MAX_TRADES
        'GetItem = Shop(EditorIndex).TradeItem(i).GetItem
        'GetValue = Shop(EditorIndex).TradeItem(i).GetValue
        'GiveItem = Shop(EditorIndex).TradeItem(i).GiveItem
        'GiveValue = Shop(EditorIndex).TradeItem(i).GiveValue
        
        'If GetItem > 0 And GiveItem > 0 Then
            'frmEditShop.lstShopTradeItem.AddItem i & ": " & GiveValue & " " & Trim(Item(GiveItem).Name) & " for " & GetValue & " " & Trim(Item(GetItem).Name)
        'Else
            'frmEditShop.lstShopTradeItem.AddItem "Empty Trade Slot"
        'End If
    'Next i
    'frmEditShop.lstShopTradeItem.ListIndex = 0
'End Sub

Public Sub UpdateShopTrade()
Dim i As Long, f As Long
Dim TradeDescription As String
Dim GiveItem(1 To MAX_GIVE_ITEMS) As Byte
Dim GiveValue(1 To MAX_GIVE_ITEMS) As Byte
Dim GetItem(1 To MAX_GET_ITEMS) As Byte
Dim GetValue(1 To MAX_GET_VALUE) As Byte

    frmEditShop.lstShopTradeItem.Clear
    For i = 1 To MAX_TRADES
        For f = 1 To MAX_GIVE_ITEMS
            GiveItem(f) = Shop(EditorIndex).TradeItem(i).GiveItem(f)
            GiveValue(f) = Shop(EditorIndex).TradeItem(i).GiveValue(f)
        Next f
        For f = 1 To MAX_GET_ITEMS
            GetItem(f) = Shop(EditorIndex).TradeItem(i).GetItem(f)
            GetValue(f) = Shop(EditorIndex).TradeItem(i).GetValue(f)
        Next f
            
        TradeDescription = ""
        
        For f = 1 To MAX_GET_ITEMS
            If (GetItem(f) > 0) And (GetValue(f) > 0) Then
                TradeDescription = TradeDescription & GetValue(f) & " x " & Trim(Item(GetItem(f)).Name) & " + "
            End If
        Next f
        If TradeDescription = "" Then
            frmEditShop.lstShopTradeItem.AddItem "Empty Trade Slot"
            frmEditShop.lstShopTradeItem.ListIndex = 0
            Exit Sub
        End If
        TradeDescription = Left(TradeDescription, (Len(TradeDescription) - 2))
        TradeDescription = TradeDescription & "for "
        For f = 1 To MAX_GIVE_ITEMS
            If (GiveItem(f) > 0) And (GiveValue(f) > 0) Then
                TradeDescription = TradeDescription & GiveValue(f) & " x " & Trim(Item(GiveItem(f)).Name) & " + "
            End If
        Next f
        If TradeDescription = "" Then
            frmEditShop.lstShopTradeItem.AddItem i & "Empty Trade Slot"
            frmEditShop.lstShopTradeItem.ListIndex = 0
            Exit Sub
        End If
        TradeDescription = Left(TradeDescription, (Len(TradeDescription) - 2))
        frmEditShop.lstShopTradeItem.AddItem i & ": " & TradeDescription
        frmEditShop.lstShopTradeItem.ListIndex = 0
    Next i
End Sub

Public Sub ShopEditorOk()
    Shop(EditorIndex).Name = frmEditShop.txtShopName.Text
    Shop(EditorIndex).FixesItems = frmEditShop.chkShopFixesItems.Value
    
    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmEditShop
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmEditShop
End Sub

Public Sub SpellEditorInit()
Dim i As Long
    frmCClient.picSpellsBack.Picture = LoadPicture(App.Path & GFX_PATH & "Spells" & GFX_EXT)
    frmCClient.tmrSpellPic.Enabled = True

    frmCClient.cmbSpellClassReq.Clear
    frmCClient.cmbSpellClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmCClient.cmbSpellClassReq.AddItem Trim(Class(i).Name)
    Next i
    
    frmCClient.txtSpellName.Text = Trim(Spell(EditorIndex).Name)
    frmCClient.scrlSpellPic.Max = MAX_SPELLS
    frmCClient.scrlSpellPic.Value = Spell(EditorIndex).SpellSprite
    frmCClient.cmbSpellClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmCClient.scrlSpellLevelReq.Value = Spell(EditorIndex).LevelReq
        
    frmCClient.cmbSpellType.ListIndex = Spell(EditorIndex).Type
    If Spell(EditorIndex).Type = SPELL_TYPE_STAT Then
        frmCClient.fraSpellStats.Visible = True
        frmCClient.fraSpellGiveItem.Visible = False
        frmCClient.cmbSpellStat.ListIndex = Spell(EditorIndex).Data1
        frmCClient.txtSpellStatMod.Text = Spell(EditorIndex).Data2
    Else
        frmCClient.fraSpellStats.Visible = False
        frmCClient.fraSpellGiveItem.Visible = True
        frmCClient.scrlSpellItemnum.Max = MAX_ITEMS
        frmCClient.scrlSpellItemnum.Value = Spell(EditorIndex).Data1
        frmCClient.scrlSpellItemValue.Value = Spell(EditorIndex).Data2
    End If
        
    frmCClient.picEditSpell.Visible = True
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).Name = frmCClient.txtSpellName.Text
    Spell(EditorIndex).SpellSprite = frmCClient.scrlSpellPic.Value
    Spell(EditorIndex).ClassReq = frmCClient.cmbSpellClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmCClient.scrlSpellLevelReq.Value
    
    Spell(EditorIndex).Type = frmCClient.cmbSpellType.ListIndex
    If Spell(EditorIndex).Type = SPELL_TYPE_STAT Then
        Spell(EditorIndex).Data1 = frmCClient.cmbSpellStat.ListIndex
        Spell(EditorIndex).Data2 = Val(frmCClient.txtSpellStatMod.Text)
    Else
        Spell(EditorIndex).Data1 = frmCClient.scrlSpellItemnum.Value
        Spell(EditorIndex).Data2 = frmCClient.scrlSpellItemValue.Value
    End If
    Spell(EditorIndex).Data3 = 0
    
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    frmCClient.tmrSpellPic.Enabled = False
    frmCClient.picEditSpell.Visible = False
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    frmCClient.tmrSpellPic.Enabled = False
    frmCClient.picEditSpell.Visible = False
End Sub

Public Sub SpellEditorBltItem()
    Call BitBlt(frmCClient.picSpellPic.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picSpellsBack.hdc, 4 * PIC_X, frmCClient.scrlSpellPic.Value * PIC_Y, SRCCOPY)
End Sub

Public Sub SkillEditorInit()
Dim i As Long
    frmCClient.picSkillsBack.Picture = LoadPicture(App.Path & GFX_PATH & "Skills" & GFX_EXT)
    frmCClient.tmrSkillPic.Enabled = True

    frmCClient.cmbSkillClassReq.Clear
    frmCClient.cmbSkillClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmCClient.cmbSkillClassReq.AddItem Trim(Class(i).Name)
    Next i
    
    frmCClient.txtSkillName.Text = Trim(Skill(EditorIndex).Name)
    frmCClient.scrlSkillPic.Max = MAX_SKILLS
    frmCClient.scrlSkillPic.Value = Skill(EditorIndex).SkillSprite
    frmCClient.cmbSkillClassReq.ListIndex = Skill(EditorIndex).ClassReq
    frmCClient.scrlSkillLevelReq.Value = Skill(EditorIndex).LevelReq
        
    frmCClient.cmbSkillType.ListIndex = Skill(EditorIndex).Type
    If Skill(EditorIndex).Type = SKILL_TYPE_ATTRIBUTE Then
        frmCClient.fraSkillVitals.Visible = True
        frmCClient.cmbSkillAttribute.ListIndex = Skill(EditorIndex).Data1
        frmCClient.cmbSkillChance.ListIndex = 0
        frmCClient.txtSkillAttributeGain.Text = Skill(EditorIndex).Data2
        If frmCClient.cmbSkillAttribute.ListIndex = SKILL_ATTRIBUTE_STR Then
            frmCClient.cmbWepType.ListIndex = Skill(EditorIndex).Data3
        Else
            frmCClient.cmbWepType.ListIndex = 0
        End If
    Else
        frmCClient.fraSkillVitals.Visible = False
    End If
    If Skill(EditorIndex).Type = SKILL_TYPE_CHANCE Then
        frmCClient.fraSkillChance.Visible = True
        If Skill(EditorIndex).Data1 > 0 Then
            frmCClient.cmbSkillChance.ListIndex = (Skill(EditorIndex).Data1 - 5)
        Else
            frmCClient.cmbSkillChance.ListIndex = 0
        End If
        frmCClient.txtSkillChanceGain.Text = Skill(EditorIndex).Data2
    Else
        frmCClient.fraSkillChance.Visible = False
    End If
        
    frmCClient.picEditSkill.Visible = True
End Sub

Public Sub SkillEditorOk()
    Skill(EditorIndex).Name = frmCClient.txtSkillName.Text
    Skill(EditorIndex).SkillSprite = frmCClient.scrlSkillPic.Value
    Skill(EditorIndex).ClassReq = frmCClient.cmbSkillClassReq.ListIndex
    Skill(EditorIndex).LevelReq = frmCClient.scrlSkillLevelReq.Value
    Skill(EditorIndex).Type = frmCClient.cmbSkillType.ListIndex
    If Skill(EditorIndex).Type = SKILL_TYPE_ATTRIBUTE Then
        Skill(EditorIndex).Data1 = frmCClient.cmbSkillAttribute.ListIndex
        Skill(EditorIndex).Data2 = Val(frmCClient.txtSkillAttributeGain.Text)
        Skill(EditorIndex).Data3 = frmCClient.cmbWepType.ListIndex
    ElseIf Skill(EditorIndex).Type = SKILL_TYPE_CHANCE Then
        If frmCClient.cmbSkillChance.ListIndex > 0 Then
            Skill(EditorIndex).Data1 = (frmCClient.cmbSkillChance.ListIndex + 5)
        Else
            Skill(EditorIndex).Data1 = 0
        End If
        Skill(EditorIndex).Data2 = Val(frmCClient.txtSkillChanceGain.Text)
        Skill(EditorIndex).Data3 = 0
    Else
        Skill(EditorIndex).Data1 = 0
        Skill(EditorIndex).Data2 = 0
        Skill(EditorIndex).Data3 = 0
    End If
    
    Call SendSaveSkill(EditorIndex)
    InSkillEditor = False
    frmCClient.tmrSkillPic.Enabled = False
    frmCClient.picEditSkill.Visible = False
End Sub

Public Sub SkillEditorCancel()
    InSkillEditor = False
    frmCClient.tmrSkillPic.Enabled = False
    frmCClient.picEditSkill.Visible = False
End Sub

Public Sub SkillEditorBltItem()
    Call BitBlt(frmCClient.picSkillPic.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picSkillsBack.hdc, 4 * PIC_X, frmCClient.scrlSkillPic.Value * PIC_Y, SRCCOPY)
End Sub

Public Sub QuestEditorInit()
Dim i As Long
    
    frmCClient.cmbQuestClass.Clear
    frmCClient.cmbQuestClass.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmCClient.cmbQuestClass.AddItem Trim(Class(i).Name)
    Next i
    
    frmCClient.txtQuestName.Text = Trim(Quest(EditorIndex).Name)
    frmCClient.txtQuestDescription.Text = Trim(Quest(EditorIndex).Description)
    frmCClient.cmbQuestClass.ListIndex = Quest(EditorIndex).ClassReq
    frmCClient.scrlQuestLevelMin.Value = Quest(EditorIndex).LevelMin
    frmCClient.scrlQuestLevelMax.Value = Quest(EditorIndex).LevelMax
    frmCClient.scrlQuestReward.Value = Quest(EditorIndex).Reward
    frmCClient.txtQuestRewardVal.Text = STR(Quest(EditorIndex).RewardValue)
    
    frmCClient.cmbQuestType.ListIndex = Quest(EditorIndex).Type
    If Quest(EditorIndex).Type = QUEST_TYPE_KILL Then
        frmCClient.fraKillQuest.Visible = True
        frmCClient.scrlQuestKillNpc.Value = Quest(EditorIndex).Data1
        frmCClient.scrlQuestKillQuantity.Value = Quest(EditorIndex).Data2
    Else
        frmCClient.fraKillQuest.Visible = False
    End If
    If Quest(EditorIndex).Type = QUEST_TYPE_FETCH Then
        frmCClient.fraFetchQuest.Visible = True
        frmCClient.scrlQuestFetchItem.Value = Quest(EditorIndex).Data1
        frmCClient.scrlQuestFetchQuantity.Value = Quest(EditorIndex).Data2
    Else
        frmCClient.fraFetchQuest.Visible = False
    End If
    If Quest(EditorIndex).Type = QUEST_TYPE_TRADE Then
        frmCClient.fraTradeQuest.Visible = True
    Else
        frmCClient.fraTradeQuest.Visible = False
    End If
    
    frmCClient.picEditQuest.Visible = True
End Sub

Public Sub QuestEditorOk()
    Quest(EditorIndex).Name = frmCClient.txtQuestName.Text
    Quest(EditorIndex).Description = frmCClient.txtQuestDescription.Text
    Quest(EditorIndex).SetBy = frmCClient.cmbQuestType.ListIndex
    Quest(EditorIndex).ClassReq = frmCClient.cmbQuestClass.ListIndex
    Quest(EditorIndex).LevelMin = frmCClient.scrlQuestLevelMin.Value
    Quest(EditorIndex).LevelMax = frmCClient.scrlQuestLevelMax.Value
    Quest(EditorIndex).Reward = frmCClient.scrlQuestReward.Value
    Quest(EditorIndex).RewardValue = Val(frmCClient.txtQuestRewardVal.Text)
    Quest(EditorIndex).Type = frmCClient.cmbQuestType.ListIndex
    If Quest(EditorIndex).Type = QUEST_TYPE_KILL Then
        Quest(EditorIndex).Data1 = frmCClient.scrlQuestKillNpc.Value
        Quest(EditorIndex).Data2 = frmCClient.scrlQuestKillQuantity.Value
        Quest(EditorIndex).Data3 = 0
    ElseIf Quest(EditorIndex).Type = QUEST_TYPE_FETCH Then
        Quest(EditorIndex).Data1 = frmCClient.scrlQuestFetchItem.Value
        Quest(EditorIndex).Data2 = frmCClient.scrlQuestFetchQuantity.Value
        Quest(EditorIndex).Data3 = 0
    ElseIf Quest(EditorIndex).Type = QUEST_TYPE_TRADE Then
        Quest(EditorIndex).Data1 = 0
        Quest(EditorIndex).Data2 = 0
        Quest(EditorIndex).Data3 = 0
    Else
        Quest(EditorIndex).Data1 = 0
        Quest(EditorIndex).Data2 = 0
        Quest(EditorIndex).Data3 = 0
    End If
    
    Call SendSaveQuest(EditorIndex)
    InQuestEditor = False
    frmCClient.picEditQuest.Visible = False
End Sub

Public Sub QuestEditorCancel()
    InQuestEditor = False
    frmCClient.picEditQuest.Visible = False
End Sub

Public Sub GUIEditorInit()
Dim i As Long
    
    frmDesign.txtDesignName.Text = Trim(GUI(EditorIndex).Name)
    frmDesign.txtDesignDesigner.Text = Trim(GUI(EditorIndex).Designer)
    frmDesign.lblDesignPicNum.Caption = 0
    For i = 1 To 7
        If GUI(EditorIndex).Background(i).Data5 > 0 Then
            frmDesign.picDesignBackground(i).Picture = LoadPicture(App.Path & DESIGN_PATH & i & "\" & GUI(EditorIndex).Background(i).Data5 & GFX_EXT)
        Else
            frmDesign.picDesignBackground(i).Picture = LoadPicture()
        End If
        frmDesign.picDesignBackground(i).Left = GUI(EditorIndex).Background(i).Data1
        frmDesign.picDesignBackground(i).Top = GUI(EditorIndex).Background(i).Data2
        frmDesign.picDesignBackground(i).Height = GUI(EditorIndex).Background(i).Data3
        frmDesign.picDesignBackground(i).Width = GUI(EditorIndex).Background(i).Data4
    Next i
    For i = 1 To 5
        frmDesign.lblDesignMenu(i).Left = GUI(EditorIndex).Menu(i).Data1
        frmDesign.lblDesignMenu(i).Top = GUI(EditorIndex).Menu(i).Data2
        frmDesign.lblDesignMenu(i).Height = GUI(EditorIndex).Menu(i).Data3
        frmDesign.lblDesignMenu(i).Width = GUI(EditorIndex).Menu(i).Data4
    Next i
    For i = 1 To 4
        frmDesign.lblDesignLogin(i).Left = GUI(EditorIndex).Login(i).Data1
        frmDesign.lblDesignLogin(i).Top = GUI(EditorIndex).Login(i).Data2
        frmDesign.lblDesignLogin(i).Height = GUI(EditorIndex).Login(i).Data3
        frmDesign.lblDesignLogin(i).Width = GUI(EditorIndex).Login(i).Data4
    Next i
    For i = 1 To 4
        frmDesign.lblDesignNewAcc(i).Left = GUI(EditorIndex).NewAcc(i).Data1
        frmDesign.lblDesignNewAcc(i).Top = GUI(EditorIndex).NewAcc(i).Data2
        frmDesign.lblDesignNewAcc(i).Height = GUI(EditorIndex).NewAcc(i).Data3
        frmDesign.lblDesignNewAcc(i).Width = GUI(EditorIndex).NewAcc(i).Data4
    Next i
    For i = 1 To 4
        frmDesign.lblDesignDelAcc(i).Left = GUI(EditorIndex).DelAcc(i).Data1
        frmDesign.lblDesignDelAcc(i).Top = GUI(EditorIndex).DelAcc(i).Data2
        frmDesign.lblDesignDelAcc(i).Height = GUI(EditorIndex).DelAcc(i).Data3
        frmDesign.lblDesignDelAcc(i).Width = GUI(EditorIndex).DelAcc(i).Data4
    Next i
    For i = 1 To 2
        frmDesign.lblDesignCredits(i).Left = GUI(EditorIndex).Credits(i).Data1
        frmDesign.lblDesignCredits(i).Top = GUI(EditorIndex).Credits(i).Data2
        frmDesign.lblDesignCredits(i).Height = GUI(EditorIndex).Credits(i).Data3
        frmDesign.lblDesignCredits(i).Width = GUI(EditorIndex).Credits(i).Data4
    Next i
    For i = 1 To 5
        frmDesign.lblDesignChars(i).Left = GUI(EditorIndex).Chars(i).Data1
        frmDesign.lblDesignChars(i).Top = GUI(EditorIndex).Chars(i).Data2
        frmDesign.lblDesignChars(i).Height = GUI(EditorIndex).Chars(i).Data3
        frmDesign.lblDesignChars(i).Width = GUI(EditorIndex).Chars(i).Data4
    Next i
    For i = 1 To 14
        frmDesign.lblDesignNewChar(i).Left = GUI(EditorIndex).NewChar(i).Data1
        frmDesign.lblDesignNewChar(i).Top = GUI(EditorIndex).NewChar(i).Data2
        frmDesign.lblDesignNewChar(i).Height = GUI(EditorIndex).NewChar(i).Data3
        frmDesign.lblDesignNewChar(i).Width = GUI(EditorIndex).NewChar(i).Data4
    Next i
    
    frmDesign.Show vbModal
End Sub

Public Sub GUIEditorSend()
    GUI(EditorIndex).Name = frmDesign.txtDesignName.Text
    GUI(EditorIndex).Designer = frmDesign.txtDesignDesigner.Text
    
    Call SendSaveGUI(EditorIndex)
    InGUIEditor = False
    Unload frmDesign
End Sub

Public Sub GUIEditorCancel()
    InGUIEditor = False
    Unload frmDesign
End Sub

'Sub ResizeGUI()
    'If frmCClient.WindowState <> vbMinimized Then
        'frmCClient.txtChat.Height = Int(frmCClient.Height / Screen.TwipsPerPixelY) - frmCClient.txtChat.top - 32
        'frmCClient.txtChat.Width = Int(frmCClient.Width / Screen.TwipsPerPixelX) - 8
    'End If
'End Sub

'Sub PlayerSearch(Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim x1 As Long, y1 As Long

    'x1 = Int(x / PIC_X)
    'y1 = Int(y / PIC_Y)
    
    'If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        'Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    'End If
'End Sub

Sub ClearTempTile()
Dim x As Long, y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            TempTile(x, y).DoorOpen = NO
        Next x
    Next y
End Sub

Sub ClearPushTile()
Dim x As Long, y As Long

    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            PushTile(x, y).Pushed = 0
        Next x
    Next y
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
    Player(Index).SPEED = 0
    Player(Index).MAGI = 0
    Player(Index).DEX = 0
        
    For n = 1 To MAX_INV
        Player(Index).Inv(n).Num = 0
        Player(Index).Inv(n).Value = 0
        Player(Index).Inv(n).Dur = 0
    Next n
        
    Player(Index).ArmorSlot = 0
    Player(Index).WeaponSlot = 0
    Player(Index).HelmetSlot = 0
    Player(Index).ShieldSlot = 0
    Player(Index).AmuletSlot = 0
    Player(Index).RingSlot = 0
    Player(Index).ArrowSlot = 0
        
    Player(Index).Map = 0
    Player(Index).x = 0
    Player(Index).y = 0
    Player(Index).Dir = 0
    
    ' Client use only
    Player(Index).MaxHp = 0
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

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim(Player(Index).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
End Sub

'Function GetPlayerPK(ByVal Index As Long) As Long
    'GetPlayerPK = Player(Index).PK
'End Function

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
    GetPlayerMaxHP = Player(Index).MaxHp
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
    GetPlayerSPEED = Player(Index).SPEED
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal SPEED As Long)
    Player(Index).SPEED = SPEED
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    GetPlayerMAGI = Player(Index).MAGI
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal MAGI As Long)
    Player(Index).MAGI = MAGI
End Sub

Function GetPlayerDEX(ByVal Index As Long) As Long
    GetPlayerDEX = Player(Index).DEX
End Function

Sub SetPlayerDEX(ByVal Index As Long, ByVal DEX As Long)
    Player(Index).DEX = DEX
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).x
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal x As Long)
    Player(Index).x = x
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal y As Long)
    Player(Index).y = y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub
