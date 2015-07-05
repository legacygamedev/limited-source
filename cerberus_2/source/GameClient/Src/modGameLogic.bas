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
        frmIPconfig.Show vbModal
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
    'InEditor = False
    'InItemsEditor = False
    'InNpcEditor = False
    'InShopEditor = False
    
    QuestNpcNum = 0
    
    ' Clear out players
    'For I = 1 To MAX_PLAYERS
        'Call ClearPlayer(I)
    'Next I
    'Call ClearTempTile
    'Call ClearPushTile
    
    ' Check and load preferences
    Call CheckPrefs
    Call LoadPrefs
    
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
    Call SavePrefs
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
        
        If frmCClient.picPlayerInventory.Visible = True Then
        ' Visual Inventory
        Dim Q As Long
        Dim Qq As Long
        
            For Q = 0 To MAX_INV - 1
                Qq = Player(MyIndex).Inv(Q + 1).Num
               
                If frmCClient.picPlayerInv(Q).Picture <> LoadPicture() Then
                frmCClient.picPlayerInv(Q).Picture = LoadPicture()
                Else
                    If Qq = 0 Then
                        frmCClient.picPlayerInv(Q).Picture = LoadPicture()
                    Else
                        'Call BitBlt(frmCClient.picPlayerInv(Q).hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, (Item(Qq).pic - Int(Item(Qq).pic / 6) * 6) * PIC_X, Int(Item(Qq).pic / 6) * PIC_Y, SRCCOPY)
                        Call BitBlt(frmCClient.picPlayerInv(Q).hdc, 2, 2, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, 0, Item(Qq).Pic * PIC_Y, SRCCOPY)
                    End If
                End If
            Next Q
        End If
                
        ' Blit out tiles layers ground/anim1/anim2
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
        
        ' Blit out the resources
        For i = 1 To MAX_MAP_RESOURCES
            Call BltResource(i)
        Next i
        
        ' Blit out the npcs
        For i = 1 To MAX_MAP_NPCS
            Call BltNpc(i)
        Next i
        
        ' Blit out arrows
        For i = 1 To HighIndex
             If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                 Call BltArrow(i)
             End If
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
        
        If GettingMap = False Then
            ' Blit player message overhead
            If GetTickCount < MessageTime + 2000 Then
                Call DrawText(TexthDC, Player(MyIndex).x * PIC_X + sx - (Int(LenB(MsgMessage)) / 2) * 3 + Player(MyIndex).XOffset, Player(MyIndex).y * PIC_Y + sx - 22 + Player(MyIndex).YOffset - iv, MsgMessage, QBColor(MessageColor))
                iv = iv + 1
            End If
            ' Blit player warning overhead
            If GetTickCount < WarnMsgTime + 2000 Then
                Call DrawText(TexthDC, Player(MyIndex).x * PIC_X + sx - (Int(LenB(WarnMessage)) / 2) * 3 + Player(MyIndex).XOffset, Player(MyIndex).y * PIC_Y + sx - 22 + Player(MyIndex).YOffset - vii, WarnMessage, QBColor(WarnMsgColor))
                vii = vii + 1
            End If
            ' Blit player damage figures overhead
            If GetTickCount < NPCDmgTime + 2000 Then
                Call DrawText(TexthDC, Player(MyIndex).x * PIC_X + sx + 8 + (Int(Len(NPCDmgDamage)) / 2) * 3 + Player(MyIndex).XOffset, Player(MyIndex).y * PIC_Y + sx - 22 + Player(MyIndex).YOffset - ii, NPCDmgDamage, QBColor(NPCDmgColor))
                ii = ii + 1
            End If
            ' Blit npc damage figures overhead
            If NPCWho > 0 Then
                If MapNpc(NPCWho).Num > 0 Then
                    If GetTickCount < DmgTime + 2000 Then
                        Call DrawText(TexthDC, MapNpc(NPCWho).x * PIC_X + sx + 8 + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).XOffset, MapNpc(NPCWho).y * PIC_Y + sx - 22 + MapNpc(NPCWho).YOffset - iii, DmgDamage, QBColor(DmgColor))
                    End If
                    iii = iii + 1
                End If
            End If
            ' Blit npc message overhead
            If NpcMsgWho > 0 Then
                If MapNpc(NpcMsgWho).Num > 0 Then
                    If GetTickCount < NpcMessageTime + 2000 Then
                        Call DrawText(TexthDC, MapNpc(NpcMsgWho).x * PIC_X + sx - (Int(LenB(NpcMsgMessage)) / 2) * 3 + MapNpc(NpcMsgWho).XOffset, MapNpc(NpcMsgWho).y * PIC_Y + sx - 22 + MapNpc(NpcMsgWho).YOffset - vi, NpcMsgMessage, QBColor(NpcMessageColor))
                    End If
                    vi = vi + 1
                End If
            End If
            ' Blit pk message overhead
            If GetTickCount < PKMsgTime + 2000 Then
                Call DrawText(TexthDC, Player(PKMsgWho).x * PIC_X + sx - (Int(LenB(PKMsgMessage)) / 2) * 3 + Player(PKMsgWho).XOffset, Player(PKMsgWho).y * PIC_Y + sx - 22 + Player(PKMsgWho).YOffset - viv, PKMsgMessage, QBColor(PKMsgColor))
                viv = viv + 1
            End If
            ' Blit player killer damage figures overhead
            If GetTickCount < PKDmgTime + 2000 Then
                Call DrawText(TexthDC, Player(PKWho).x * PIC_X + sx + 8 + (Int(Len(PKDmgDamage)) / 2) * 3 + Player(PKWho).XOffset, Player(PKWho).y * PIC_Y + sx - 22 + Player(PKWho).YOffset - viii, PKDmgDamage, QBColor(PKDmgColor))
                viii = viii + 1
            End If
            ' Blit Resource message overhead
            If ResourceWho > 0 Then
                If MapResource(ResourceWho).Num > 0 Then
                    If GetTickCount < ResourceMsgTime + 2000 Then
                        Call DrawText(TexthDC, MapResource(ResourceWho).x * PIC_X + sx - (Int(LenB(ResourceMsgMessage)) / 2) * 3, MapResource(ResourceWho).y * PIC_Y + sx - 22 - xi, ResourceMsgMessage, QBColor(ResourceMsgColor))
                    End If
                    xi = xi + 1
                End If
            End If
            ' Blit resource damage figures overhead
            If ResourceDmgWho > 0 Then
                If MapResource(ResourceDmgWho).Num > 0 Then
                    If GetTickCount < ResourceDmgTime + 2000 Then
                        Call DrawText(TexthDC, MapResource(ResourceDmgWho).x * PIC_X + sx + 8 + (Int(Len(ResourceDmgDamage)) / 2) * 3, MapResource(ResourceDmgWho).y * PIC_Y + sx - 22 - xii, ResourceDmgDamage, QBColor(ResourceDmgColor))
                    End If
                    xii = xii + 1
                End If
            End If
            ' Blit Item message overhead
            If ItemWho > 0 Then
                If MapItem(ItemWho).Num > 0 Then
                    If GetTickCount < ItemMsgTime + 2000 Then
                        Call DrawText(TexthDC, MapItem(ItemWho).x * PIC_X + sx - (Int(LenB(ItemMsgMessage)) / 2) * 3, MapItem(ItemWho).y * PIC_Y + sx - 22 - xiv, ItemMsgMessage, QBColor(ItemMsgColor))
                    End If
                    xiv = xiv + 1
                End If
            End If
        End If
        
        'For i = 1 To HighIndex
            'If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                'Call BltPlayerName(i)
            'End If
        'Next i
                
        '' Blit out attribs if in editor
        'If InEditor Then
            'For y = 0 To MAX_MAPY
                'For x = 0 To MAX_MAPX
                    'With Map.Tile(x, y)
                        'If .Type = TILE_TYPE_BLOCKED Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "B", QBColor(BrightRed))
                        'If .Type = TILE_TYPE_WARP Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "W", QBColor(BrightBlue))
                        'If .Type = TILE_TYPE_ITEM Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "I", QBColor(White))
                        'If .Type = TILE_TYPE_NPCAVOID Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "N", QBColor(White))
                        'If .Type = TILE_TYPE_KEY Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "K", QBColor(White))
                        'If .Type = TILE_TYPE_KEYOPEN Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "O", QBColor(White))
                        'If .Type = TILE_TYPE_PUSHBLOCK Then Call DrawText(TexthDC, x * PIC_X + 8, y * PIC_Y + 8, "P", QBColor(White))
                    'End With
                'Next x
            'Next y
        'End If
        
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
        
        ' Get the rect for the back buffer to blit from
        With rec
            .top = 0
            .Bottom = (MAX_MAPY + 1) * PIC_Y
            .Left = 0
            .Right = (MAX_MAPX + 1) * PIC_X
        End With
        
        ' Get the rect to blit to
        Call DX.GetWindowRect(frmCClient.picScreen.hWnd, rec_pos)
        With rec_pos
            .Bottom = .top + ((MAX_MAPY + 1) * PIC_Y)
            .Right = .Left + ((MAX_MAPX + 1) * PIC_X)
        End With
        
        ' Blit the backbuffer
        Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec, DDBLT_WAIT)
        
        ' Check if player is trying to move
        Call CheckMovement
        
        '' Check to see if player is trying to attack
        Call CheckAttack
        
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
        .top = y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    With rec
        .top = Int(Ground / 14) * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = (Ground - Int(Ground / 14) * 14) * PIC_X
        .Right = .Left + PIC_X
    End With
    Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT)
    
    If Mask > 0 And TempTile(x, y).DoorOpen = NO And PushTile(x, y).Pushed = NO And PushTile(x, y).Moving = 0 Then
        With rec
            .top = Int(Mask / 14) * PIC_Y
            .Bottom = .top + PIC_Y
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
             rec.top = Int(Mask2 / 14) * PIC_Y
             rec.Bottom = rec.top + PIC_Y
             rec.Left = (Mask2 - Int(Mask2 / 14) * 14) * PIC_X
             rec.Right = rec.Left + PIC_X
             'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
             Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim > 0 Then
             rec.top = Int(Anim / 14) * PIC_Y
             rec.Bottom = rec.top + PIC_Y
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
        .top = Int(Map.Tile(x, y).Mask / 14) * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = (Map.Tile(x, y).Mask - Int(Map.Tile(x, y).Mask / 14) * 14) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    x1 = x * PIC_X + PushTile(x, y).XOffset
    y1 = y * PIC_Y + PushTile(x, y).YOffset
    
    ' Check if its out of bounds because of the offset
    If y1 < 0 Then
        y1 = 0
        With rec
            .top = .top + (y1 * -1)
        End With
    End If
        
    Call DD_BackBuffer.BltFast(x1, y1, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Public Sub BltItem(ByVal ItemNum As Long)

    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .top = MapItem(ItemNum).y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = MapItem(ItemNum).x * PIC_X
        .Right = .Left + PIC_X
    End With

    With rec
        .top = Item(MapItem(ItemNum).Num).Pic * PIC_Y
        .Bottom = .top + PIC_Y
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
        .top = y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Fringe = Map.Tile(x, y).Fringe
    Fringe2 = Map.Tile(x, y).Fringe2
    FAnim = Map.Tile(x, y).FAnim
        
    If Fringe > 0 Then
        rec.top = Int(Fringe / 14) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (Fringe - Int(Fringe / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If

    If (MapAnim = 0) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        If Fringe2 > 0 Then
            rec.top = Int(Fringe2 / 14) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Fringe2 - Int(Fringe2 / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            'Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(x * PIC_X, y * PIC_Y, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If FAnim > 0 Then
            rec.top = Int(FAnim / 14) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
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
        .top = y * PIC_Y
        .Bottom = .top + PIC_Y
        .Left = x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    Light = Map.Tile(x, y).Light
        
    If Light > 0 Then
        With rec
            .top = Int(Light / 14) * PIC_Y
            .Bottom = .top + PIC_Y
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
        .top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
        .Bottom = .top + PIC_Y
        .Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
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
    
    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .top = (((GetPlayerSprite(Index) * 2) * PIC_Y) + PIC_Y)
        .Bottom = .top + PIC_Y
        .Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
        .Right = .Left + PIC_X
    End With
    
    x = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
    y = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        With rec
            .top = .top + (y * -1)
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
    
    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .top = ((GetPlayerSprite(Index) * 2) * PIC_Y)
        .Bottom = .top + PIC_Y
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
            .top = .top - y
            y = 0
        End With
    End If
        
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltArrow(ByVal Index As Long)
Dim x As Long, y As Long, i As Long, z As Long
'Dim BX As Long, BY As Long

For z = 1 To MAX_PLAYER_ARROWS
    If Player(Index).Arrow(z).Arrow > 0 Then
    
        rec.top = Player(Index).Arrow(z).ArrowAnim * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = Player(Index).Arrow(z).ArrowPosition * PIC_X
        rec.Right = rec.Left + PIC_X
       
        If GetTickCount > Player(Index).Arrow(z).ArrowTime + 30 Then
             Player(Index).Arrow(z).ArrowTime = GetTickCount
             Player(Index).Arrow(z).ArrowVarX = Player(Index).Arrow(z).ArrowVarX + 10
             Player(Index).Arrow(z).ArrowVarY = Player(Index).Arrow(z).ArrowVarY + 10
        End If
       
        If Player(Index).Arrow(z).ArrowPosition = 0 Then
             x = Player(Index).Arrow(z).ArrowX
             y = Player(Index).Arrow(z).ArrowY + Int(Player(Index).Arrow(z).ArrowVarY / 32)
             'If Y > Player(Index).Arrow(z).ArrowY + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
             If y > Player(Index).Arrow(z).ArrowY + Item(GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index))).Data2 - 2 Then
                Player(Index).Arrow(z).Arrow = 0
                If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                    Call SendData("lastarrow" & SEP_CHAR & GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
             End If
            
             If y <= MAX_MAPY Then
                 'Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset + Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                 Call DD_BackBuffer.BltFast(Player(Index).Arrow(z).ArrowX * PIC_X, Player(Index).Arrow(z).ArrowY * PIC_Y + Player(Index).Arrow(z).ArrowVarY, DD_ArrowSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
        End If
       
        If Player(Index).Arrow(z).ArrowPosition = 1 Then
             x = Player(Index).Arrow(z).ArrowX
             y = Player(Index).Arrow(z).ArrowY - Int(Player(Index).Arrow(z).ArrowVarY / 32)
             'If Y < Player(Index).Arrow(z).ArrowY - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
             If y < Player(Index).Arrow(z).ArrowY - Item(GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index))).Data2 + 2 Then
                 Player(Index).Arrow(z).Arrow = 0
                 If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                    Call SendData("lastarrow" & SEP_CHAR & GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) & SEP_CHAR & END_CHAR)
                    Exit Sub
                 End If
             End If
            
             If y >= 0 Then
                 'Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset - Player(Index).Arrow(z).ArrowVarY, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                 Call DD_BackBuffer.BltFast(Player(Index).Arrow(z).ArrowX * PIC_X, Player(Index).Arrow(z).ArrowY * PIC_Y - Player(Index).Arrow(z).ArrowVarY, DD_ArrowSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
        End If
       
        If Player(Index).Arrow(z).ArrowPosition = 2 Then
             x = Player(Index).Arrow(z).ArrowX + Int(Player(Index).Arrow(z).ArrowVarX / 32)
             y = Player(Index).Arrow(z).ArrowY
             'If X > Player(Index).Arrow(z).ArrowX + Arrows(Player(Index).Arrow(z).ArrowNum).Range - 2 Then
             If x > Player(Index).Arrow(z).ArrowX + Item(GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index))).Data2 - 2 Then
                 Player(Index).Arrow(z).Arrow = 0
                 If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                    Call SendData("lastarrow" & SEP_CHAR & GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) & SEP_CHAR & END_CHAR)
                    Exit Sub
                 End If
             End If
            
             If x <= MAX_MAPX Then
                 'Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset + Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                 Call DD_BackBuffer.BltFast(Player(Index).Arrow(z).ArrowX * PIC_X + Player(Index).Arrow(z).ArrowVarX, Player(Index).Arrow(z).ArrowY * PIC_Y, DD_ArrowSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
        End If
       
        If Player(Index).Arrow(z).ArrowPosition = 3 Then
             x = Player(Index).Arrow(z).ArrowX - Int(Player(Index).Arrow(z).ArrowVarX / 32)
             y = Player(Index).Arrow(z).ArrowY
             'If X < Player(Index).Arrow(z).ArrowX - Arrows(Player(Index).Arrow(z).ArrowNum).Range + 2 Then
             If x < Player(Index).Arrow(z).ArrowX - Item(GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index))).Data2 + 2 Then
                 Player(Index).Arrow(z).Arrow = 0
                 If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                    Call SendData("lastarrow" & SEP_CHAR & GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) & SEP_CHAR & END_CHAR)
                    Exit Sub
                 End If
             End If
            
             If x >= 0 Then
                 'Call DD_BackBuffer.BltFast((Player(Index).Arrow(z).ArrowX - NewPlayerX) * PIC_X + sx - NewXOffset - Player(Index).Arrow(z).ArrowVarX, (Player(Index).Arrow(z).ArrowY - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ArrowAnim, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                 Call DD_BackBuffer.BltFast(Player(Index).Arrow(z).ArrowX * PIC_X - Player(Index).Arrow(z).ArrowVarX, Player(Index).Arrow(z).ArrowY * PIC_Y, DD_ArrowSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
             End If
        End If
       
        If x >= 0 And x <= MAX_MAPX Then
             If y >= 0 And y <= MAX_MAPY Then
                 If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Then
                     Player(Index).Arrow(z).Arrow = 0
                     If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                        Call SendData("lastarrow" & SEP_CHAR & GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) & SEP_CHAR & END_CHAR)
                        Exit Sub
                     End If
                 End If
             End If
        End If
       
        For i = 1 To HighIndex
           If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                 If GetPlayerX(i) = x And GetPlayerY(i) = y Then
                     If Index = MyIndex Then
                           Call SendData("arrowhit" & SEP_CHAR & 0 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                     End If
                     If Index <> i Then Player(Index).Arrow(z).Arrow = 0
                     If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                         Call SendData("lastarrow" & SEP_CHAR & GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) & SEP_CHAR & END_CHAR)
                         Exit Sub
                     End If
                     Exit Sub
                 End If
             End If
        Next i
       
        For i = 1 To MAX_MAP_NPCS
             If MapNpc(i).Num > 0 Then
                 If MapNpc(i).x = x And MapNpc(i).y = y Then
                     If Index = MyIndex Then
                           Call SendData("arrowhit" & SEP_CHAR & 1 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                     End If
                     Player(Index).Arrow(z).Arrow = 0
                     If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                         Call SendData("lastarrow" & SEP_CHAR & GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) & SEP_CHAR & END_CHAR)
                         Exit Sub
                     End If
                     Exit Sub
                 End If
             End If
        Next i
        
        For i = 1 To MAX_MAP_RESOURCES
             If MapResource(i).Num > 0 Then
                 If MapResource(i).x = x And MapResource(i).y = y Then
                     If Index = MyIndex Then
                           Call SendData("arrowhit" & SEP_CHAR & 2 & SEP_CHAR & i & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR)
                     End If
                     Player(Index).Arrow(z).Arrow = 0
                     If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                         Call SendData("lastarrow" & SEP_CHAR & GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)) & SEP_CHAR & END_CHAR)
                         Exit Sub
                     End If
                     Exit Sub
                 End If
             End If
        Next i
        
        'For BX = 0 To MAX_MAPX
        '   For BY = 0 To MAX_MAPY
        '        If Map(GetPlayerMap(MyIndex)).Tile(BX, BY).Type = TILE_TYPE_NPC_SPAWN Then
        '              For i = 1 To MAX_ATTRIBUTE_NPCS
         '                 If MapAttributeNpc(i, BX, BY).X = X And MapAttributeNpc(i, BX, BY).Y = Y Then
         '                     If Index = MyIndex Then
          '                         Call SendData("arrowhit" & SEP_CHAR & 2 & SEP_CHAR & i & SEP_CHAR & X & SEP_CHAR & Y & SEP_CHAR & BX & SEP_CHAR & BY & SEP_CHAR & END_CHAR)
          '                    End If
          '                    Player(Index).Arrow(z).Arrow = 0
          '                    Exit Sub
          '               End If
                     'Next i
          '      End If
        '    Next BY
        'Next BX
    End If
Next z
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
    With MapNpc(MapNpcNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .top = (((Npc(MapNpc(MapNpcNum).Num).Sprite * 2) * PIC_Y) + PIC_Y)
        .Bottom = .top + PIC_Y
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
            .top = .top + (y * -1)
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
    With MapNpc(MapNpcNum)
        If .AttackTimer + 1000 < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With
    
    With rec
        .top = ((Npc(MapNpc(MapNpcNum).Num).Sprite * 2) * PIC_Y)
        .Bottom = .top + PIC_Y
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
            .top = .top - y
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
        .top = MapResource(MapResourceNum).y * PIC_Y '+ MapResource(MapResourceNum).YOffset
        .Bottom = .top + PIC_Y
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
            .top = (((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y) + PIC_Y)
            .Bottom = .top + PIC_Y
            .Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + 64
        End With
    ElseIf Npc(MapResource(MapResourceNum).Num).Big = 2 Then
        With rec
            .top = (((Npc(MapResource(MapResourceNum).Num).Sprite * 4) * PIC_Y) + (3 * PIC_Y))
            .Bottom = .top + PIC_Y
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
        .top = (((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y) + PIC_Y)
        .Bottom = .top + PIC_Y
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
            .top = .top + (y * -1)
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
        .top = MapResource(MapResourceNum).y * PIC_Y '+ MapResource(MapResourceNum).YOffset
        .Bottom = .top + PIC_Y
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
            .top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y)
            .Bottom = .top + PIC_Y
            .Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + 64
        End With
    ElseIf Npc(MapResource(MapResourceNum).Num).Big = 2 Then
        With rec
            .top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 4) * PIC_Y)
            .Bottom = .top + (4 * PIC_Y)
            .Left = 0 '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + PIC_X
        End With
        With rec1
            .top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 4) * PIC_Y)
            .Bottom = .top + (3 * PIC_Y)
            .Left = PIC_X '(MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            .Right = .Left + PIC_X
        End With
        With rec2
            .top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 4) * PIC_Y)
            .Bottom = .top + (4 * PIC_Y)
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
        .top = ((Npc(MapResource(MapResourceNum).Num).Sprite * 2) * PIC_Y)
        .Bottom = .top + PIC_Y
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
            .top = .top - y
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
            Call AddText("Available Commands: /help, /inv, /stats, /quests, /skills, /spells", HelpColor) ', /info, /who, /fps, /inv, /stats, /train, /trade, /party, /join, /leave", HelpColor)
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
        
        '' Whos Online
        'If LCase(Mid(MyText, 1, 4)) = "/who" Then
            'Call SendWhosOnline
            'MyText = ""
            'Exit Sub
        'End If
                        
        '' Checking fps
        'If LCase(Mid(MyText, 1, 4)) = "/fps" Then
            'Call AddText("FPS: " & GameFPS, Pink)
            'MyText = ""
            'Exit Sub
        'End If
                
        ' Show inventory
        If LCase(Mid(MyText, 1, 4)) = "/inv" Then
            If frmCClient.chkPlayerInventoryPin.Value = False Then
                Call UpdateInventory
                Call UpdateVisInventory
                frmCClient.chkPlayerInventoryPin.Value = Checked
                frmCClient.picPlayerInventory.Visible = True
                frmCClient.tmrPlayerInventory.Enabled = False
            Else
                frmCClient.chkPlayerInventoryPin.Value = Unchecked
                frmCClient.picPlayerInventory.Visible = False
                frmCClient.tmrPlayerInventory.Enabled = True
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Show stats
        If LCase(Mid(MyText, 1, 6)) = "/stats" Then
            If frmCClient.chkPlayerStatsPin.Value = False Then
                frmCClient.chkPlayerStatsPin.Value = Checked
                frmCClient.picPlayerStats.Visible = True
                frmCClient.tmrPlayerStats.Enabled = False
            Else
                frmCClient.chkPlayerStatsPin.Value = Unchecked
                frmCClient.picPlayerStats.Visible = False
                frmCClient.tmrPlayerStats.Enabled = True
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Show quests
        If LCase(Mid(MyText, 1, 7)) = "/quests" Then
            If frmCClient.chkPlayerQuestPin.Value = False Then
                frmCClient.chkPlayerQuestPin.Value = Checked
                frmCClient.picPlayerQuests.Visible = True
                frmCClient.tmrPlayerQuests.Enabled = False
            Else
                frmCClient.chkPlayerQuestPin.Value = Unchecked
                frmCClient.picPlayerQuests.Visible = False
                frmCClient.tmrPlayerQuests.Enabled = True
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Show quests
        If LCase(Mid(MyText, 1, 7)) = "/skills" Then
            If frmCClient.chkPlayerSkillsPin.Value = False Then
                frmCClient.chkPlayerSkillsPin.Value = Checked
                frmCClient.picPlayerSkills.Visible = True
                frmCClient.tmrPlayerSkills.Enabled = False
            Else
                frmCClient.chkPlayerSkillsPin.Value = Unchecked
                frmCClient.picPlayerSkills.Visible = False
                frmCClient.tmrPlayerSkills.Enabled = True
            End If
            MyText = ""
            Exit Sub
        End If
        
        ' Show quests
        If LCase(Mid(MyText, 1, 7)) = "/spells" Then
            If frmCClient.chkPlayerSpellsPin.Value = False Then
                frmCClient.chkPlayerSpellsPin.Value = Checked
                frmCClient.picPlayerSpells.Visible = True
                frmCClient.tmrPlayerSpells.Enabled = False
            Else
                frmCClient.chkPlayerSpellsPin.Value = Unchecked
                frmCClient.picPlayerSpells.Visible = False
                frmCClient.tmrPlayerSpells.Enabled = True
            End If
            MyText = ""
            Exit Sub
        End If
    
        '' Show training
        'If LCase(Mid(MyText, 1, 6)) = "/train" Then
            'frmTraining.Show vbModal
            'MyText = ""
            'Exit Sub
        'End If

        '' Request stats
        'If LCase(Mid(MyText, 1, 6)) = "/trade" Then
            'Call SendData("trade" & SEP_CHAR & END_CHAR)
            'MyText = ""
            'Exit Sub
        'End If
        
        '' Party request
        'If LCase(Mid(MyText, 1, 6)) = "/party" Then
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

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then 'And Trim(MyText) = "" Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & SEP_CHAR & END_CHAR)
    End If
End Sub

Public Sub CheckAttack()
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
                    Call SendPlayerDir
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
                        If d <> DIR_UP Then
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
                        If d <> DIR_UP Then
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

Public Sub UpdateInventory()
Dim i As Long

    frmCClient.lstPlayerInventory.Clear
    
    ' Show the inventory
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                frmCClient.lstPlayerInventory.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
            Else
                ' Check if this item is being worn
                If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerArrowSlot(MyIndex) = i Then
                    frmCClient.lstPlayerInventory.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                Else
                    frmCClient.lstPlayerInventory.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                End If
            End If
        Else
            frmCClient.lstPlayerInventory.AddItem "<free inventory slot>"
        End If
    Next i
    
    frmCClient.lstPlayerInventory.ListIndex = 0
End Sub

Public Sub UpdateVisInventory()
Dim Index As Long
Dim d As Long

    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) <> Index Then frmCClient.picPlayerShield.Picture = LoadPicture()
        If GetPlayerWeaponSlot(MyIndex) <> Index Then frmCClient.picPlayerWeapon.Picture = LoadPicture()
        If GetPlayerHelmetSlot(MyIndex) <> Index Then frmCClient.picPlayerHelmet.Picture = LoadPicture()
        If GetPlayerArmorSlot(MyIndex) <> Index Then frmCClient.picPlayerArmour.Picture = LoadPicture()
        If GetPlayerAmuletSlot(MyIndex) <> Index Then frmCClient.picPlayerAmulet.Picture = LoadPicture()
        If GetPlayerRingSlot(MyIndex) <> Index Then frmCClient.picPlayerRing.Picture = LoadPicture()
        If GetPlayerArrowSlot(MyIndex) <> Index Then
            frmCClient.picPlayerArrows.Picture = LoadPicture()
            frmCClient.lblArrows.Caption = 0
        End If
    Next Index
    
    For Index = 1 To MAX_INV
        'If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerShield.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        'If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerWeapon.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        'If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerHelmet.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        'If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerArmour.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, (Item(GetPlayerInvItemNum(MyIndex, Index)).pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerShield.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, Index)).Pic * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerWeapon.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, Index)).Pic * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerHelmet.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, Index)).Pic * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerArmour.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, Index)).Pic * PIC_Y, SRCCOPY)
        If GetPlayerAmuletSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerAmulet.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, Index)).Pic * PIC_Y, SRCCOPY)
        If GetPlayerRingSlot(MyIndex) = Index Then Call BitBlt(frmCClient.picPlayerRing.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, Index)).Pic * PIC_Y, SRCCOPY)
        If GetPlayerArrowSlot(MyIndex) = Index Then
            Call BitBlt(frmCClient.picPlayerArrows.hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picInventoryItems.hdc, 0, Item(GetPlayerInvItemNum(MyIndex, Index)).Pic * PIC_Y, SRCCOPY)
            frmCClient.lblArrows.Caption = GetPlayerInvItemDur(MyIndex, GetPlayerArrowSlot(MyIndex))
        End If
    Next Index
    
    frmCClient.shpSelectedItem.top = frmCClient.picPlayerInv(frmCClient.lstPlayerInventory.ListIndex).top - 1
    frmCClient.shpSelectedItem.Left = frmCClient.picPlayerInv(frmCClient.lstPlayerInventory.ListIndex).Left - 1
    
    frmCClient.shpEquiped(0).Visible = False
    frmCClient.shpEquiped(1).Visible = False
    frmCClient.shpEquiped(2).Visible = False
    frmCClient.shpEquiped(3).Visible = False
    frmCClient.shpEquiped(4).Visible = False
    frmCClient.shpEquiped(5).Visible = False
    frmCClient.shpEquiped(6).Visible = False

    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).Num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type <> ITEM_TYPE_CURRENCY Then
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    frmCClient.shpEquiped(0).Visible = True
                    frmCClient.shpEquiped(0).top = frmCClient.picPlayerInv(d).top - 2
                    frmCClient.shpEquiped(0).Left = frmCClient.picPlayerInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    frmCClient.shpEquiped(1).Visible = True
                    frmCClient.shpEquiped(1).top = frmCClient.picPlayerInv(d).top - 2
                    frmCClient.shpEquiped(1).Left = frmCClient.picPlayerInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    frmCClient.shpEquiped(2).Visible = True
                    frmCClient.shpEquiped(2).top = frmCClient.picPlayerInv(d).top - 2
                    frmCClient.shpEquiped(2).Left = frmCClient.picPlayerInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    frmCClient.shpEquiped(3).Visible = True
                    frmCClient.shpEquiped(3).top = frmCClient.picPlayerInv(d).top - 2
                    frmCClient.shpEquiped(3).Left = frmCClient.picPlayerInv(d).Left - 2
                ElseIf GetPlayerAmuletSlot(MyIndex) = d + 1 Then
                    frmCClient.shpEquiped(4).Visible = True
                    frmCClient.shpEquiped(4).top = frmCClient.picPlayerInv(d).top - 2
                    frmCClient.shpEquiped(4).Left = frmCClient.picPlayerInv(d).Left - 2
                ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                    frmCClient.shpEquiped(5).Visible = True
                    frmCClient.shpEquiped(5).top = frmCClient.picPlayerInv(d).top - 2
                    frmCClient.shpEquiped(5).Left = frmCClient.picPlayerInv(d).Left - 2
                ElseIf GetPlayerArrowSlot(MyIndex) = d + 1 Then
                    frmCClient.shpEquiped(6).Visible = True
                    frmCClient.shpEquiped(6).top = frmCClient.picPlayerInv(d).top - 2
                    frmCClient.shpEquiped(6).Left = frmCClient.picPlayerInv(d).Left - 2
                End If
            End If
        End If
    Next d
    If frmCClient.picScreen.Visible Then frmCClient.picScreen.SetFocus
End Sub

Public Sub UpdateVisSpells()
Dim Index As Long
Dim d As Long

    For Index = 1 To MAX_PLAYER_SPELLS
        For d = 1 To MAX_ITEMS
            If Player(MyIndex).Spells(Index).Num <> d Then frmCClient.picSpell(Index - 1).Picture = LoadPicture()
        Next d
    Next Index
    
    For Index = 1 To MAX_PLAYER_SPELLS
        For d = 1 To MAX_ITEMS
            If Player(MyIndex).Spells(Index).Num = d Then Call BitBlt(frmCClient.picSpell(Index - 1).hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picSpellSprite.hdc, 4 * PIC_X, Spell(Player(MyIndex).Spells(Index).Num).SpellSprite * PIC_Y, SRCCOPY)
            frmCClient.picSpell(Index - 1).Refresh
        Next d
    Next Index
End Sub

Public Sub UpdateVisSkills()
Dim Index As Long
Dim d As Long

    For Index = 1 To MAX_PLAYER_SKILLS
        For d = 1 To MAX_ITEMS
            If Player(MyIndex).Skills(Index).Num <> d Then frmCClient.picSkill(Index - 1).Picture = LoadPicture()
            'frmCClient.picSkill(Index - 1).Refresh
        Next d
    Next Index
    
    For Index = 1 To MAX_PLAYER_SKILLS
        For d = 1 To MAX_ITEMS
            If Player(MyIndex).Skills(Index).Num = d Then Call BitBlt(frmCClient.picSkill(Index - 1).hdc, 0, 0, PIC_X, PIC_Y, frmCClient.picSkillSprite.hdc, 4 * PIC_X, Skill(Player(MyIndex).Skills(Index).Num).SkillSprite * PIC_Y, SRCCOPY)
            frmCClient.picSkill(Index - 1).Refresh
        Next d
    Next Index
End Sub

'Sub ResizeGUI()
    'If frmCClient.WindowState <> vbMinimized Then
        'frmCClient.txtChat.Height = Int(frmCClient.Height / Screen.TwipsPerPixelY) - frmCClient.txtChat.top - 32
        'frmCClient.txtChat.Width = Int(frmCClient.Width / Screen.TwipsPerPixelX) - 8
    'End If
'End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1 As Long, y1 As Long
Dim i As Long

    x1 = Int(x / PIC_X)
    y1 = Int(y / PIC_Y)
    
    'For i = MAX_MAP_ITEMS To 1
        'If (MapItem(i).x = x1) And (MapItem(i).y = y1) Then
            'ItemWho = i
            'ItemMsgTime = GetTickCount
            'ItemMsgColor = Yellow
            'ItemMsgMessage = "Item : " & Trim(Item(MapItem(i).Num).Name)
            'xiv = 0
        'End If
    'Next i
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
End Sub

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
    
    For n = 1 To MAX_PLAYER_QUESTS
        Player(Index).Quests(n).Num = 0
        Player(Index).Quests(n).SetMap = 0
        Player(Index).Quests(n).SetBy = 0
        Player(Index).Quests(n).Amount = 0
        Player(Index).Quests(n).Count = 0
    Next n
    
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

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).Num = ItemNum
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

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Spells(SpellSlot).Num
End Function

Function GetPlayerSpellLevel(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpellLevel = Player(Index).Spells(SpellSlot).Level
End Function

Function GetPlayerSpellExp(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpellExp = Player(Index).Spells(SpellSlot).EXP
End Function

Function GetPlayerSkill(ByVal Index As Long, ByVal SkillSlot As Long) As Long
    GetPlayerSkill = Player(Index).Skills(SkillSlot).Num
End Function

Function GetPlayerSkillLevel(ByVal Index As Long, ByVal SkillSlot As Long) As Long
    GetPlayerSkillLevel = Player(Index).Skills(SkillSlot).Level
End Function

Function GetPlayerSkillExp(ByVal Index As Long, ByVal SkillSlot As Long) As Long
    GetPlayerSkillExp = Player(Index).Skills(SkillSlot).EXP
End Function

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

Function GetPlayerAmuletSlot(ByVal Index As Long) As Long
    GetPlayerAmuletSlot = Player(Index).AmuletSlot
End Function

Sub SetPlayerAmuletSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).AmuletSlot = InvNum
End Sub

Function GetPlayerRingSlot(ByVal Index As Long) As Long
    GetPlayerRingSlot = Player(Index).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).RingSlot = InvNum
End Sub

Function GetPlayerArrowSlot(ByVal Index As Long) As Long
    GetPlayerArrowSlot = Player(Index).ArrowSlot
End Function

Sub SetPlayerArrowSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).ArrowSlot = InvNum
End Sub
