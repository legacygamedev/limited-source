Attribute VB_Name = "modGameLogic"
Option Explicit

Public Function TwipsToPixels(lngTwips As Long, _
        lngDirection As Long) As Long

    ' Handle to device
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

    ' Handle to device
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
    DoEvents
End Sub

Sub MenuState(ByVal State As Long)
    Connected = True

    frmSendGetData.Visible = True

    Call SetStatus("Connecting to Server...")

    Select Case State
        Case MENU_STATE_NEWACCOUNT
            frmNewAccount.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected! Creating Account...")

                If Not frmNewAccount.txtEmail.Visible Then
                    Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text, "NOMAIL")
                Else
                    Call SendNewAccount(frmNewAccount.txtName.Text, frmNewAccount.txtPassword.Text, frmNewAccount.txtEmail.Text)
                End If
            End If

        Case MENU_STATE_DELACCOUNT
            frmDeleteAccount.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Deleting Account...")

                Call SendDelAccount(frmDeleteAccount.txtName.Text, frmDeleteAccount.txtPassword.Text)
            End If

        Case MENU_STATE_LOGIN
            frmLogin.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Logging In...")

                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If

        Case MENU_STATE_AUTO_LOGIN
            frmMainMenu.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Logging In...")

                Call SendLogin(frmLogin.txtName.Text, frmLogin.txtPassword.Text)
            End If

        Case MENU_STATE_NEWCHAR
            frmChars.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Receiving Classes...")

                If SpriteSize = 1 Then
                    frmNewChar.Picture4.top = frmNewChar.Picture4.top - 32
                    frmNewChar.Picture4.Height = 69
                    frmNewChar.picPic.Height = 65
                End If

                If CustomPlayers <> 0 Then
                    frmNewChar.HScroll1.Visible = True
                    frmNewChar.HScroll2.Visible = True
                    frmNewChar.HScroll3.Visible = True
                End If

                Call SendGetClasses
            End If

        Case MENU_STATE_ADDCHAR
            frmNewChar.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Creating Character...")

                If frmNewChar.optMale.value Then
                    Call SendAddChar(frmNewChar.txtName, 0, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1, frmNewChar.HScroll1.value, frmNewChar.HScroll2.value, frmNewChar.HScroll3.value)
                Else
                    Call SendAddChar(frmNewChar.txtName, 1, frmNewChar.cmbClass.ListIndex, frmChars.lstChars.ListIndex + 1, frmNewChar.HScroll1.value, frmNewChar.HScroll2.value, frmNewChar.HScroll3.value)
                End If
            End If

        Case MENU_STATE_DELCHAR
            frmChars.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Deleting Character...")

                Call SendDelChar(frmChars.lstChars.ListIndex + 1)
            End If

        Case MENU_STATE_USECHAR
            frmChars.Visible = False

            If ConnectToServer Then
                Call SetStatus("Connected. Entering " & GAME_NAME & "...")

                Call SendUseChar(frmChars.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected And Connected = True Then
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True

        Call MsgBox("The server is currently offline. Please try to reconnect in a few minutes or visit " & WEBSITE, vbOKOnly, GAME_NAME)
    End If
End Sub

Sub GameInit()
    Call InitDirectX
    Call StopBGM

    InGame = True

    ' Check for divide by 0 error
    If GetPlayerMaxHP(MyIndex) > 0 Then
        frmStable.shpHP.Width = ((GetPlayerHP(MyIndex)) / (GetPlayerMaxHP(MyIndex))) * 150
        frmStable.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
    End If

    ' Check for divide by 0 error
    If GetPlayerMaxMP(MyIndex) > 0 Then
        frmStable.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
        frmStable.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
    End If
    
    ' Unload main menu forms after character logs in.
    Unload frmSendGetData
    Unload frmMainMenu
    Unload frmChars
    Unload frmNewChar
    Unload frmSendGetData
    
    frmStable.Picture = LoadPicture(App.Path & "\GUI\800X600.jpg")
    frmStable.picCharStatus.Picture = LoadPicture(App.Path & "\GUI\minimenus.jpg")
    frmStable.picEquipment.Picture = LoadPicture(App.Path & "\GUI\minimenus.jpg")
    frmStable.picPlayerSpells.Picture = LoadPicture(App.Path & "\GUI\minimenus.jpg")
    frmStable.picInventory.Picture = LoadPicture(App.Path & "\GUI\minimenus.jpg")
    frmStable.picGuildAdmin.Picture = LoadPicture(App.Path & "\GUI\minimenus.jpg")
    frmStable.picWhosOnline.Picture = LoadPicture(App.Path & "\GUI\minimenus.jpg")
    frmStable.picGuildMember.Picture = LoadPicture(App.Path & "\GUI\minimenus.jpg")
    frmStable.picInventory3.Picture = LoadPicture(App.Path & "\GUI\minimenus.jpg")

    frmStable.Visible = True

    On Error Resume Next

    ' Set the focus To the main form since only focussed objects may Set the focus
    frmStable.SetFocus

    frmStable.picScreen.SetFocus
End Sub

Sub GameLoop()
    Dim Tick As Long
    Dim TickFPS As Long
    Dim FPS As Long
    Dim X As Long
    Dim Y As Long
    Dim I As Long
    Dim z As Long

    ' This will be re-enabled once Eclipse Evolution 2.7 is released. [Mellowz]
    On Error Resume Next

    ' Used for calculating fps
    TickFPS = GetTickCount
    FPS = 0

    ' *******************************************
    ' * ECLIPSE EVOLUTION MAIN GAME LOOP BEGIN  *
    ' *******************************************
    Do While InGame
        Tick = GetTickCount

        If frmStable.WindowState = 0 Then

            ' Check if we need to restore surfaces
            If NeedToRestoreSurfaces Then
                DD.RestoreAllSurfaces
                Call InitSurfaces
            End If

            If Not GettingMap Then

                ' Check to make sure they aren't trying to auto do anything
                If GetAsyncKeyState(VK_UP) >= 0 And DirUp Then
                    DirUp = False
                End If
                If GetAsyncKeyState(VK_DOWN) >= 0 And DirDown Then
                    DirDown = False
                End If
                If GetAsyncKeyState(VK_LEFT) >= 0 And DirLeft Then
                    DirLeft = False
                End If
                If GetAsyncKeyState(VK_RIGHT) >= 0 And DirRight Then
                    DirRight = False
                End If
                If GetAsyncKeyState(VK_CONTROL) >= 0 And ControlDown Then
                    ControlDown = False
                End If
                If GetAsyncKeyState(VK_SHIFT) >= 0 And ShiftDown Then
                    ShiftDown = False
                End If
    
                ' Check to make sure we are still connected
                If Not IsConnected Then
                    InGame = False
                    Exit Do
                End If

                ' Visual Inventory
                Dim Q As Long
                Dim Qq As Long
                Dim IT As Long

                If GetTickCount > IT + 500 And frmStable.picInventory.Visible = True Then
                    For Q = 0 To MAX_INV - 1
                        Qq = Player(MyIndex).Inv(Q + 1).Num

                        If frmStable.picInv(Q).Picture <> LoadPicture() Then
                            frmStable.picInv(Q).Picture = LoadPicture()
                        Else
                            If Qq = 0 Then
                                frmStable.picInv(Q).Picture = LoadPicture()
                            Else
                                Call BitBlt(frmStable.picInv(Q).hDC, 0, 0, PIC_X, PIC_Y, frmStable.picItems.hDC, (Item(Qq).Pic - Int(Item(Qq).Pic / 6) * 6) * PIC_X, Int(Item(Qq).Pic / 6) * PIC_Y, SRCCOPY)
                            End If
                        End If
                    Next Q
                End If
                
                            
                NewX = 12
                NewY = 9

                NewPlayerY = Player(MyIndex).Y - NewY
                NewPlayerX = Player(MyIndex).X - NewX

                NewX = NewX * PIC_X
                NewY = NewY * PIC_Y

                NewXOffset = Player(MyIndex).xOffset
                NewYOffset = Player(MyIndex).yOffset

                If Player(MyIndex).Y - 9 < 1 Then
                    NewY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = 0
                    If Player(MyIndex).Y = 9 And Player(MyIndex).Dir = DIR_UP Then
                        NewPlayerY = Player(MyIndex).Y - 9
                        NewY = 9 * PIC_Y
                        NewYOffset = Player(MyIndex).yOffset
                    End If
                ElseIf Player(MyIndex).Y + 11 > MAX_MAPY + 1 Then
                    NewY = (Player(MyIndex).Y - (MAX_MAPY - 18)) * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = MAX_MAPY - 18
                    If Player(MyIndex).Y = MAX_MAPY - 9 And Player(MyIndex).Dir = DIR_DOWN Then
                        NewPlayerY = Player(MyIndex).Y - 9
                        NewY = 9 * PIC_Y
                        NewYOffset = Player(MyIndex).yOffset
                    End If
                End If

                If Player(MyIndex).X - 12 < 1 Then
                    NewX = Player(MyIndex).X * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = 0
                    If Player(MyIndex).X = 12 And Player(MyIndex).Dir = DIR_LEFT Then
                        NewPlayerX = Player(MyIndex).X - 12
                        NewX = 12 * PIC_X
                        NewXOffset = Player(MyIndex).xOffset
                    End If
                ElseIf Player(MyIndex).X + 14 > MAX_MAPX + 1 Then
                    NewX = (Player(MyIndex).X - (MAX_MAPX - 24)) * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = MAX_MAPX - 24
                    If Player(MyIndex).X = MAX_MAPX - 12 And Player(MyIndex).Dir = DIR_RIGHT Then
                        NewPlayerX = Player(MyIndex).X - 12
                        NewX = 12 * PIC_X
                        NewXOffset = Player(MyIndex).xOffset
                    End If
                End If

                ScreenX = GetScreenLeft(MyIndex)
                ScreenY = GetScreenTop(MyIndex)
                ScreenX2 = GetScreenRight(MyIndex)
                ScreenY2 = GetScreenBottom(MyIndex)

                If ScreenX < 0 Then
                    ScreenX = 0
                    ScreenX2 = 25
                ElseIf ScreenX2 > MAX_MAPX Then
                    ScreenX2 = MAX_MAPX
                    ScreenX = MAX_MAPX - 25
                End If
           
                If ScreenY < 0 Then
                    ScreenY = 0
                    ScreenY2 = 19
                ElseIf ScreenY2 > MAX_MAPY Then
                    ScreenY2 = MAX_MAPY
                    ScreenY = MAX_MAPY - 19
                End If

                sx = 32
                If MAX_MAPX = 19 Then
                    NewX = Player(MyIndex).X * PIC_X + Player(MyIndex).xOffset
                    NewXOffset = 0
                    NewPlayerX = 0
                    NewY = Player(MyIndex).Y * PIC_Y + Player(MyIndex).yOffset
                    NewYOffset = 0
                    NewPlayerY = 0
                    ScreenX = 0
                    ScreenY = 0
                    ScreenX2 = MAX_MAPX
                    ScreenY2 = MAX_MAPY
                    sx = 0
                End If

                ' Blit out tiles layers ground/anim1/anim2
                For Y = ScreenY To ScreenY2
                    For X = ScreenX To ScreenX2
                        Call BltTile(X, Y)
                    Next X
                Next Y

                If ScreenMode = 0 Then
                
                    ' Blit out the items
                    For I = 1 To MAX_MAP_ITEMS
                        If MapItem(I).Num > 0 Then
                            Call BltItem(I)
                        End If
                    Next I
                    
                    ' Blit out NPC hp bars
                    If frmStable.chkNpcBar.value = Checked Then
                        For I = 1 To MAX_MAP_NPCS
                            Call BltNpcBars(I)
                        Next I
                    End If
                    
                     ' Blit players bar
                    If frmStable.chkPlayerBar.value = Checked Then
                        For I = 1 To MAX_PLAYERS
                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                    Call BltPlayerBars(I)
                                End If
                            End If
                        Next I
                    End If

                    ' Blit out the sprite change attribute
                    If Right$(Trim$(Map(GetPlayerMap(MyIndex)).name), 1) = "*" Then
                        For Y = ScreenY To ScreenY2
                            For X = ScreenX To ScreenX2
                                Call BltSpriteChange(X, Y)
                            Next X
                        Next Y
                    End If

                    ' Blit out grapple
                    For I = 1 To MAX_PLAYERS
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                Call Bltgrapple(I)
                            End If
                        End If
                    Next I

                    ' Blit out players and arrows
                    For I = 1 To MAX_PLAYERS
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                Call BltPlayer(I)
                                Call BltArrow(I)
                            End If
                        End If
                    Next I
                    
                    ' Blit out the npc base
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(I).Num > 0 Then
                            Call BltNpcBody(I)
                        End If
                    Next I

                    ' Blit out the npc tops
                    For I = 1 To MAX_MAP_NPCS
                        If MapNpc(I).Num > 0 Then
                            Call BltNpcTop(I)
                        End If
                    Next I

                    ' Blit out players top
                    If SpriteSize >= 1 Then
                        For I = 1 To MAX_PLAYERS
                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                    Call BltPlayerTop(I)
                                End If
                            End If
                        Next I
                    End If

                    ' Blt out the spells
                    For I = 1 To MAX_PLAYERS
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                Call BltSpell(I)
                            End If
                        End If
                    Next I

                    ' Blt out the scripted spells
                    For I = 1 To MAX_SCRIPTSPELLS
                        If ScriptSpell(I).SpellNum > 0 Then
                            If ScriptSpell(I).SpellNum <= MAX_SPELLS Then
                                If ScriptSpell(I).CastedSpell = YES Then
                                    Call BltScriptSpell(I)
                                End If
                            End If
                        End If
                    Next I
                    
                    ' Draw 'level up!' text
                    For I = 1 To MAX_PLAYERS
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                Call BltLevelUp(I)
                            End If
                        End If
                    Next I

                End If

                ' Blit out tile layer fringe
                For Y = ScreenY To ScreenY2
                    For X = ScreenX To ScreenX2
                        Call BltFringeTile(X, Y)
                    Next X
                Next Y

                ' Check for roof tiles
                For Y = ScreenY To ScreenY2
                    For X = ScreenX To ScreenX2
                        If Not IsTileRoof(X, Y) Then
                            Call BltFringe2Tile(X, Y)
                        End If
                    Next X
                Next Y
                
                ' Blit out emoticons
                For I = 1 To MAX_PLAYERS
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                            Call BltEmoticons(I)
                        End If
                    End If
                Next I

                ' Draw night (for normal players).
                If GameTime = TIME_NIGHT Then
                    'Call AddText("Nighttime enabled", GREEN)
                    If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
                        If Not InEditor Then
                            Call Night
                        End If
                    End If
                End If
            
                ' Draw night (for administrators).
                If InEditor Then
                    If NightMode = 1 Then
                        'Call AddText("Nighttime enabled", GREEN)
                        Call Night
                    End If
                End If
            
                ' Draw weather (for all players)
                If Map(GetPlayerMap(MyIndex)).Indoors = 0 Then
                    If Map(GetPlayerMap(MyIndex)).Weather <> 0 Then
                        Call BltMapWeather
                    End If
            
                    Call BltWeather
                End If

                If InEditor Then
                    If GridMode = 1 Then
                        For Y = ScreenY To ScreenY2
                            For X = ScreenX To ScreenX2
                                Call BltTile2(X * PIC_X, Y * PIC_Y, 0)
                            Next X
                        Next Y
                    End If
                End If

                ' Lock the backbuffer so we can draw text and names
                TexthDC = DD_BackBuffer.GetDC

                If ScreenMode = 0 Then
                
                    ' Draw NPC's damage on player
                    If frmStable.chkNpcDamage.value = 1 Then
                        If frmStable.chkPlayerName.value = 0 Then
                            If GetTickCount < NPCDmgTime + 2000 Then
                                Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BRIGHTRED))
                            End If
                        Else
                            If GetPlayerGuild(MyIndex) <> vbNullString Then
                                If GetTickCount < NPCDmgTime + 2000 Then
                                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 42 - ii + sx, NPCDmgDamage, QBColor(BRIGHTRED))
                                End If
                            Else
                                If GetTickCount < NPCDmgTime + 2000 Then
                                    Call DrawText(TexthDC, (Int(Len(NPCDmgDamage)) / 2) * 3 + NewX + sx, NewY - 22 - ii + sx, NPCDmgDamage, QBColor(BRIGHTRED))
                                End If
                            End If
                        End If
                        ii = ii + 1
                    End If

                    ' Draw player's damage on NPC
                    If frmStable.chkPlayerDamage.value = 1 Then
                        If NPCWho > 0 Then
                            If MapNpc(NPCWho).Num > 0 Then
                                If frmStable.chkNpcName.value = 0 Then
                                    If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                                        If GetTickCount < DmgTime + 2000 Then
                                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 20 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(WHITE))
                                        End If
                                    Else
                                        If GetTickCount < DmgTime + 2000 Then
                                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 47 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(WHITE))
                                        End If
                                    End If
                                Else
                                    If Npc(MapNpc(NPCWho).Num).Big = 0 Then
                                        If GetTickCount < DmgTime + 2000 Then
                                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 30 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(WHITE))
                                        End If
                                    Else
                                        If GetTickCount < DmgTime + 2000 Then
                                            Call DrawText(TexthDC, (MapNpc(NPCWho).X - NewPlayerX) * PIC_X + sx + (Int(Len(DmgDamage)) / 2) * 3 + MapNpc(NPCWho).xOffset - NewXOffset, (MapNpc(NPCWho).Y - NewPlayerY) * PIC_Y + sx - 57 + MapNpc(NPCWho).yOffset - NewYOffset - iii, DmgDamage, QBColor(WHITE))
                                        End If
                                    End If
                                End If
                                iii = iii + 1
                            End If
                        End If
                    End If
                    
                    ' Draw player name and guild name
                    If frmStable.chkPlayerName.value = 1 Then
                        For I = 1 To MAX_PLAYERS
                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                    Call BltPlayerGuildName(I)
                                    Call BltPlayerName(I)
                                End If
                            End If
                        Next I
                    End If

                    ' speech bubble stuffs
                    If ReadINI("CONFIG", "SpeechBubbles", App.Path & "\config.ini") = 1 Then
                        For I = 1 To MAX_PLAYERS
                            If IsPlaying(I) Then
                                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                                    If Bubble(I).Text <> vbNullString Then
                                        Call BltPlayerText(I)
                                    End If
    
                                    If GetTickCount() > Bubble(I).Created + DISPLAY_BUBBLE_TIME Then
                                        Bubble(I).Text = vbNullString
                                    End If
                                End If
                            End If
                        Next I
                    End If

                    ' scriptbubble stuffs
                    For z = 1 To MAX_BUBBLES
                        If IsPlaying(MyIndex) Then
                            If GetPlayerMap(MyIndex) = ScriptBubble(z).Map Then
                                If ScriptBubble(z).Text <> vbNullString Then
                                    Call Bltscriptbubble(z, ScriptBubble(z).Map, ScriptBubble(z).X, ScriptBubble(z).Y, ScriptBubble(z).Colour)
                                End If
    
                                If GetTickCount() > ScriptBubble(z).Created + DISPLAY_BUBBLE_TIME Then
                                    ScriptBubble(z).Text = vbNullString
                                End If
                            End If
                        End If
                    Next z

                    ' Draw NPC Names
                    If ReadINI("CONFIG", "NPCName", App.Path & "\config.ini") = 1 Then
                        For I = LBound(MapNpc) To UBound(MapNpc)
                            If MapNpc(I).Num > 0 Then
                                Call BltMapNPCName(I)
                            End If
                        Next I
                    End If

                    ' Blit out attribs if in editor
                    If InEditor Then
                        For Y = 0 To MAX_MAPY
                            For X = 0 To MAX_MAPX
                                With Map(GetPlayerMap(MyIndex)).Tile(X, Y)
                                    If .Type = TILE_TYPE_BLOCKED Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "B", QBColor(BRIGHTRED))
                                    End If
                                    If .Type = TILE_TYPE_WARP Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "W", QBColor(BRIGHTBLUE))
                                    End If
                                    If .Type = TILE_TYPE_ITEM Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "I", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_NPCAVOID Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_KEY Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_KEYOPEN Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "O", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_HEAL Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "H", QBColor(BRIGHTGREEN))
                                    End If
                                    If .Type = TILE_TYPE_KILL Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "K", QBColor(BRIGHTRED))
                                    End If
                                    If .Type = TILE_TYPE_SHOP Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "S", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_CBLOCK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CB", QBColor(BLACK))
                                    End If
                                    If .Type = TILE_TYPE_ARENA Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "A", QBColor(BRIGHTGREEN))
                                    End If
                                    If .Type = TILE_TYPE_SOUND Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PS", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_SPRITE_CHANGE Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(GREY))
                                    End If
                                    If .Type = TILE_TYPE_SIGN Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SI", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_DOOR Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "D", QBColor(BLACK))
                                    End If
                                    If .Type = TILE_TYPE_NOTICE Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "N", QBColor(BRIGHTGREEN))
                                    End If
                                    If .Type = TILE_TYPE_CHEST Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "C", QBColor(BROWN))
                                    End If
                                    If .Type = TILE_TYPE_CLASS_CHANGE Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "CG", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_SCRIPTED Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "SC", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_HOUSE Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "PH", QBColor(YELLOW))
                                    End If
                                    If .light > 0 Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 18 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 14 - (NewPlayerY * PIC_Y) - NewYOffset, "L", QBColor(YELLOW))
                                    End If
                                    If .Type = TILE_TYPE_BANK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "BANK", QBColor(BRIGHTRED))
                                    End If
                                    If .Type = TILE_TYPE_GUILDBLOCK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "GB", QBColor(MAGENTA))
                                    End If
                                    If .Type = TILE_TYPE_HOOKSHOT Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "GS", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_WALKTHRU Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "WT", QBColor(RED))
                                    End If
                                    If .Type = TILE_TYPE_ROOF Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "RF", QBColor(RED))
                                    End If
                                    If .Type = TILE_TYPE_ROOFBLOCK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "RFB", QBColor(BRIGHTRED))
                                    End If
                                    If .Type = TILE_TYPE_ONCLICK Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "OC", QBColor(WHITE))
                                    End If
                                    If .Type = TILE_TYPE_LOWER_STAT Then
                                        Call DrawText(TexthDC, X * PIC_X + sx + 8 - (NewPlayerX * PIC_X) - NewXOffset, Y * PIC_Y + sx + 8 - (NewPlayerY * PIC_Y) - NewYOffset, "-S", QBColor(BRIGHTRED))
                                    End If
                                End With
                            Next X
                        Next Y
                    End If

                    ' draw FPS
                    If BFPS Then
                        Call DrawText(TexthDC, 18 * PIC_X + sx, sx, "FPS: " & GameFPS, QBColor(YELLOW))
                    End If

                    ' draw cursor and player X and Y locations
                    If BLoc Then
                        Call DrawText(TexthDC, 0 + sx, 0 + sx, "Cursor (X: " & CurX & "; Y: " & CurY & ")", QBColor(YELLOW))
                        Call DrawText(TexthDC, 0 + sx, 15 + sx, "Location (X: " & GetPlayerX(MyIndex) & "; Y: " & GetPlayerY(MyIndex) & ")", QBColor(YELLOW))
                        Call DrawText(TexthDC, 0 + sx, 30 + sx, "Map #" & GetPlayerMap(MyIndex), QBColor(YELLOW))
                    End If

                    ' Draw map name
                    If Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NONE Then
                        Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).name), QBColor(BRIGHTRED))
                    ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_HOUSE Then
                        Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).name), QBColor(YELLOW))
                    ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_SAFE Then
                        Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).name), QBColor(WHITE))
                    ElseIf Map(GetPlayerMap(MyIndex)).Moral = MAP_MORAL_NO_PENALTY Then
                        Call DrawText(TexthDC, Int((20.5) * PIC_X / 2) - (Int(Len(Trim$(Map(GetPlayerMap(MyIndex)).name)) / 2) * 8) + sx, 2 + sx, Trim$(Map(GetPlayerMap(MyIndex)).name), QBColor(BLACK))
                    End If

                    For I = 1 To MAX_BLT_LINE
                        If BattlePMsg(I).Index > 0 Then
                            If BattlePMsg(I).time + 7000 > GetTickCount Then
                                Call DrawText(TexthDC, 1 + sx, BattlePMsg(I).Y + frmStable.picScreen.Height - 15 + sx, Trim$(BattlePMsg(I).Msg), QBColor(BattlePMsg(I).color))
                            Else
                                BattlePMsg(I).Done = 0
                            End If
                        End If

                        If BattleMMsg(I).Index > 0 Then
                            If BattleMMsg(I).time + 7000 > GetTickCount Then
                                Call DrawText(TexthDC, (frmStable.picScreen.Width - (Len(BattleMMsg(I).Msg) * 8)) + sx, BattleMMsg(I).Y + frmStable.picScreen.Height - 15 + sx, Trim$(BattleMMsg(I).Msg), QBColor(BattleMMsg(I).color))
                            Else
                                BattleMMsg(I).Done = 0
                            End If
                        End If
                    Next I
                        
                End If
                
            Else
                ' Lock the backbuffer so we can draw text
                TexthDC = DD_BackBuffer.GetDC
                
                ' Show player that a new map is loading
                Call DrawText(TexthDC, 36, 36, "Receiving map...", QBColor(BRIGHTCYAN))
            End If

            ' Release DC
            Call DD_BackBuffer.ReleaseDC(TexthDC)

            ' Get the rect for the back buffer to blit from
            rec.top = 0
            rec.Bottom = (MAX_MAPY + 1) * PIC_Y
            rec.Left = 0
            rec.Right = (MAX_MAPX + 1) * PIC_X

            ' Get the rect to blit to
            Call DX.GetWindowRect(frmStable.picScreen.hWnd, rec_pos)
            rec_pos.Bottom = rec_pos.top - sx + ((MAX_MAPY + 1) * PIC_Y)
            rec_pos.Right = rec_pos.Left - sx + ((MAX_MAPX + 1) * PIC_X)
            rec_pos.top = rec_pos.Bottom - ((MAX_MAPY + 1) * PIC_Y)
            rec_pos.Left = rec_pos.Right - ((MAX_MAPX + 1) * PIC_X)

            ' Blit the backbuffer
            Call DD_PrimarySurf.Blt(rec_pos, DD_BackBuffer, rec, DDBLT_WAIT)

            ' Check if player is trying to move
            Call CheckMovement

            ' Check to see if player is trying to attack
            Call CheckAttack

            ' Process player movements (actually move them)
            For I = 1 To MAX_PLAYERS
                If IsPlaying(I) Then
                    Call ProcessMovement(I)
                End If
            Next I

            ' Process npc movements (actually move them)
            For I = 1 To MAX_MAP_NPCS
                If Map(GetPlayerMap(MyIndex)).Npc(I) > 0 Then
                    Call ProcessNpcMovement(I)
                End If
            Next I

        End If

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
        Do While GetTickCount < Tick + 31
            DoEvents
            Sleep 1
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

    ' MsgBox "Connection lost!"

    ' Shutdown the game
    Call GameDestroy

    Exit Sub
End Sub

' Closes the game client.
Sub GameDestroy()
    ' Unloads all TCP-related things.
    Call TcpDestroy

    ' Unloads all DirectX objects.
    Call DestroyDirectX

    'Close sound system
    Call BASS_Free

    ' Closes the VB6 application.
    End
End Sub

Sub BltTile(ByVal X As Long, ByVal Y As Long)
    Dim Ground As Long
    Dim Mask1 As Long
    Dim Anim1 As Long
    Dim Mask2 As Long
    Dim Anim2 As Long
    Dim GroundTileSet As Byte
    Dim Mask1TileSet As Byte
    Dim Anim1TileSet As Byte
    Dim Mask2TileSet As Byte
    Dim Anim2TileSet As Byte

    Ground = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Ground
    Mask1 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).mask
    Anim1 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Anim
    Mask2 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask2
    Anim2 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).M2Anim

    GroundTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).GroundSet
    Mask1TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).MaskSet
    Anim1TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).AnimSet
    Mask2TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Mask2Set
    Anim2TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).M2AnimSet

    If TileFile(GroundTileSet) = 0 Then
        Exit Sub
    End If

    rec.top = Int(Ground / TilesInSheets) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (Ground - Int(Ground / TilesInSheets) * TilesInSheets) * PIC_X
    rec.Right = rec.Left + PIC_X

    Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(GroundTileSet), rec, DDBLTFAST_WAIT)

    If MapAnim = 0 Or Anim1 = 0 Then
        If Mask1 > 0 Then
            If TileFile(Mask1TileSet) = 0 Then
                Exit Sub
            End If

            If TempTile(X, Y).DoorOpen = NO Then
                rec.top = Int(Mask1 / TilesInSheets) * PIC_Y
                rec.Bottom = rec.top + PIC_Y
                rec.Left = (Mask1 - Int(Mask1 / TilesInSheets) * TilesInSheets) * PIC_X
                rec.Right = rec.Left + PIC_X
                
                Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Mask1TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Else
        If Anim1 > 0 Then
            If TileFile(Anim1TileSet) = 0 Then
                Exit Sub
            End If

            rec.top = Int(Anim1 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Anim1 - Int(Anim1 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Anim1TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If

    If MapAnim = 0 Or Anim2 = 0 Then
        If Mask2 > 0 Then
            If TileFile(Mask2TileSet) = 0 Then
                Exit Sub
            End If

            rec.top = Int(Mask2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Mask2 - Int(Mask2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Mask2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If Anim2 > 0 Then
            If TileFile(Anim2TileSet) = 0 Then
                Exit Sub
            End If

            rec.top = Int(Anim2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Anim2 - Int(Anim2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset + sx, (Y - NewPlayerY) * PIC_Y - NewYOffset + sx, DD_TileSurf(Anim2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    rec.top = Int(Item(MapItem(ItemNum).Num).Pic / 6) * PIC_Y
    rec.Bottom = rec.top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).Num).Pic - Int(Item(MapItem(ItemNum).Num).Pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X

    Call DD_BackBuffer.BltFast((MapItem(ItemNum).X - NewPlayerX) * PIC_X + sx - NewXOffset, (MapItem(ItemNum).Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal X As Long, ByVal Y As Long)
    Dim Fringe As Long
    Dim FAnim As Long
    Dim FringeTileSet As Byte
    Dim FAnimTileSet As Byte

    Fringe = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe
    FAnim = Map(GetPlayerMap(MyIndex)).Tile(X, Y).FAnim

    FringeTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).FringeSet
    FAnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).FAnimSet

    If MapAnim = 0 Or FAnim = 0 Then
        If Fringe > 0 Then
            If TileFile(FringeTileSet) = 0 Then
                Exit Sub
            End If

            rec.top = Int(Fringe / TilesInSheets) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Fringe - Int(Fringe / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FringeTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If FAnim > 0 Then
            If TileFile(FAnimTileSet) = 0 Then
                Exit Sub
            End If

            rec.top = Int(FAnim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (FAnim - Int(FAnim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(FAnimTileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltFringe2Tile(ByVal X As Integer, ByVal Y As Integer)
    Dim Fringe2 As Long
    Dim F2Anim As Long
    Dim Fringe2TileSet As Byte
    Dim F2AnimTileSet As Byte

    Fringe2 = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe2
    F2Anim = Map(GetPlayerMap(MyIndex)).Tile(X, Y).F2Anim

    Fringe2TileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).Fringe2Set
    F2AnimTileSet = Map(GetPlayerMap(MyIndex)).Tile(X, Y).F2AnimSet

    If MapAnim = 0 Or F2Anim = 0 Then
        If Fringe2 > 0 Then
            If TileFile(Fringe2TileSet) = 0 Then
                Exit Sub
            End If

            rec.top = Int(Fringe2 / TilesInSheets) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (Fringe2 - Int(Fringe2 / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

            Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X + sx - NewXOffset, (Y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf(Fringe2TileSet), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        If F2Anim > 0 Then
            If TileFile(F2AnimTileSet) = 0 Then
                Exit Sub
            End If

            rec.top = Int(F2Anim / TilesInSheets) * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (F2Anim - Int(F2Anim / TilesInSheets) * TilesInSheets) * PIC_X
            rec.Right = rec.Left + PIC_X

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
        attack_item = Int(Player(Index).Inv(attack_weaponslot).Num)
        If attack_item > 0 Then
            AttackSpeed = 1000 'Item(attack_item).AttackSpeed
        Else
            AttackSpeed = 1000
        End If
    Else
        AttackSpeed = 1000
    End If
    
    If WalkFix = 0 Then
        ' Check for animation
        Anim = 0
        If Player(Index).Attacking = 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (Player(Index).yOffset < PIC_Y / 2) Then
                        Anim = 1
                    End If
                Case DIR_DOWN
                    If (Player(Index).yOffset < PIC_Y / 2 * -1) Then
                        Anim = 1
                    End If
                Case DIR_LEFT
                    If (Player(Index).xOffset < PIC_Y / 2) Then
                        Anim = 1
                    End If
                Case DIR_RIGHT
                    If (Player(Index).xOffset < PIC_Y / 2 * -1) Then
                        Anim = 1
                    End If
            End Select
        Else
            If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
                Anim = 2
            End If
        End If
    Else
        ' Check for animation
          Anim = 1
        If Player(Index).Attacking = 0 Then
          Select Case GetPlayerDir(Index)
            Case DIR_UP
              If (Player(Index).yOffset > 8) Then Anim = Player(Index).Step
            Case DIR_DOWN
              If (Player(Index).yOffset < -8) Then Anim = Player(Index).Step
            Case DIR_LEFT
              If (Player(Index).xOffset > 8) Then Anim = Player(Index).Step
            Case DIR_RIGHT
              If (Player(Index).xOffset < -8) Then Anim = Player(Index).Step
          End Select
        Else
          If Player(Index).AttackTimer + 1000 > GetTickCount Then
            Anim = 2
          End If
        End If
    End If

    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
        Player(Index).Attacking = 0
    '    Player(Index).AttackTimer = 0
    End If

    ' Configure what happens if theres no items there
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

    ' 32 X 64
    If SpriteSize = 1 Then

        ' 32 X 64
        If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then

            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
            rec.Right = rec.Left + 32

            If Index = MyIndex Then
                X = NewX + sx
                Y = NewY + sx

                ' PLAYER 32 X 64 IF DIR = UP
                If GetPlayerDir(MyIndex) = DIR_UP Then

                    ' PLAYER 32 X 64 BLIT SHIELD IF DIR = UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.top = Item(Player(MyIndex).Shield).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 64 BLIT WEAPON IF DIR = UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.top = Item(Player(MyIndex).Weapon).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 64 BLIT NECKLACE IF DIR = UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.top = Item(Player(MyIndex).Necklace).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                If CustomPlayers = 0 Then
                    ' PLAYER 32 X 64 BLIT SPRITE
                    rec.top = GetPlayerSprite(MyIndex) * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    ' PLAYER 32 X 64 BLIT SPRITE
                    rec.top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 64 BLIT LEGS
                If Player(MyIndex).legs > 0 Then
                    rec.top = Item(Player(MyIndex).legs).Pic * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 64 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                    rec.top = Item(Player(MyIndex).Armor).Pic * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 64 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                    rec.top = Item(Player(MyIndex).Helmet).Pic * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 64 DIR <> UP
                If GetPlayerDir(MyIndex) <> DIR_UP Then

                    ' PLAYER 32 X 64 BLIT SHIELD IF DIR <> UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.top = Item(Player(MyIndex).Shield).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 64 BLIT WEAPON IF DIR <> UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.top = Item(Player(MyIndex).Weapon).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 64 BLIT NECKLACE IF DIR <> UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.top = Item(Player(MyIndex).Necklace).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If


            ' 32 X 64 IF OTHER PLAYER
            Else

                X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
                Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset

                ' OTHER 32 X 64 IF DIR = UP
                If GetPlayerDir(Index) = DIR_UP Then

                    ' OTHER 32 X 64 BLIT SHIELD IF DIR = UP
                    If Player(Index).Shield > 0 Then
                        rec.top = Item(Player(Index).Shield).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 64 BLIT WEAPON IF DIR = UP
                    If Player(Index).Weapon > 0 Then
                        rec.top = Item(Player(Index).Weapon).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 64 BLIT NECKLACE IF DIR = UP
                    If Player(Index).Necklace > 0 Then
                        rec.top = Item(Player(Index).Necklace).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                ' OTHER 32 X 64 BLIT SPRITE
                If 0 + CustomPlayers = 0 Then
                    rec.top = GetPlayerSprite(Index) * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 64 BLIT LEGS
                If Player(Index).legs > 0 Then
                    rec.top = Item(Player(Index).legs).Pic * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 64 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                    rec.top = Item(Player(Index).Armor).Pic * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 64 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                    rec.top = Item(Player(Index).Helmet).Pic * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 64 IF DIR <> UP
                If GetPlayerDir(Index) <> DIR_UP Then

                    ' OTHER 32 X 64 BLIT SHIELD IF DIR <> UP
                    If Player(Index).Shield > 0 Then
                        rec.top = Item(Player(Index).Shield).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 64 BLIT NECKLACE IF DIR <> UP
                    If Player(Index).Necklace > 0 Then
                        rec.top = Item(Player(Index).Necklace).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' 'OTHER 32 X 64 BLIT WEAPON IF DIR <> UP
                    If Player(Index).Weapon > 0 Then
                        rec.top = Item(Player(Index).Weapon).Pic * 64 + PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

            ' END OF PAPERDOLL FOR 32 X 64
            End If

        ' IF 32 X 64 AND NO PAPERDOLL
        Else
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * 32
            rec.Right = rec.Left + 32

            ' PLAYER 32 X 64
            If Index = MyIndex Then
                X = NewX + sx
                Y = NewY + sx

                If 0 + CustomPlayers = 0 Then
                    ' PLAYER 32 X 64 BLIT SPRITE
                    rec.top = GetPlayerSprite(MyIndex) * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    ' PLAYER 32 X 64 BLIT SPRITE
                    rec.top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

            ' OTHER 32 X 64
            Else
                X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
                Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset

                ' OTHER 32 X 64 BLIT SPRITE
                If 0 + CustomPlayers = 0 Then
                    rec.top = GetPlayerSprite(Index) * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.top = Player(Index).head * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).body * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).leg * 64 + PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

        ' END OF 32 X 64
        End If

    ' 32 X 32 LOOP
    ElseIf SpriteSize = 0 Then

        rec.top = GetPlayerSprite(Index) * PIC_Y
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X

        ' 32 X 32 PLAYER
        If Index = MyIndex Then

            ' 32 X 32 PAPERDOLLED PLAYER
            If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
                X = NewX + sx
                Y = NewY + sx

                ' PLAYER 32 X 32 IF DIR = UP
                If GetPlayerDir(MyIndex) = DIR_UP Then

                    ' PLAYER 32 X 32 BLIT SHIELD IF DIR = UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.top = Item(Player(MyIndex).Shield).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 32 BLIT WEAPON IF DIR = UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.top = Item(Player(MyIndex).Weapon).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 32 BLIT NECKLACE IF DIR = UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.top = Item(Player(MyIndex).Necklace).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                If 0 + CustomPlayers = 0 Then
                    ' PLAYER 32 X 32 BLIT SPRITE
                    rec.top = GetPlayerSprite(MyIndex) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    ' PLAYER 32 X 32 BLIT SPRITE
                    rec.top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 32 BLIT LEGS
                If Player(MyIndex).legs > 0 Then
                    rec.top = Item(Player(MyIndex).legs).Pic * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 32 BLIT ARMOR
                If Player(MyIndex).Armor > 0 Then
                    rec.top = Item(Player(MyIndex).Armor).Pic * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 32 BLIT HELMET
                If Player(MyIndex).Helmet > 0 Then
                    rec.top = Item(Player(MyIndex).Helmet).Pic * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' PLAYER 32 X 32 IF DIR <> UP
                If GetPlayerDir(MyIndex) <> DIR_UP Then

                    ' PLAYER 32 X 32 BLIT SHIELD IF DIR <> UP
                    If Player(MyIndex).Shield > 0 Then
                        rec.top = Item(Player(MyIndex).Shield).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 32 BLIT WEAPON IF DIR <> UP
                    If Player(MyIndex).Weapon > 0 Then
                        rec.top = Item(Player(MyIndex).Weapon).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' PLAYER 32 X 32 BLIT NECKLACE IF DIR <> UP
                    If Player(MyIndex).Necklace > 0 Then
                        rec.top = Item(Player(MyIndex).Necklace).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

            ' 32 X 32 IF NO PAPERDOLL ON SELF BLIT JUST SPRITE
            Else
                X = NewX + sx
                Y = NewY + sx
                If 0 + CustomPlayers = 0 Then
                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X, Y, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

        ' 32 X 32 OTHER LOOP
        Else
            X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
            Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset '- 4

            ' IF OFF TOP EDGE ADJUST
            If Y < 0 Then
                rec.top = rec.top + (Y * -1)
                Y = 0
            End If

            ' 32 X 32 OTHER PAPERDOLL LOOP
            If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then

                ' OTHER 32 X 32 IF DIR = UP
                If GetPlayerDir(Index) = DIR_UP Then

                    ' OTHER 32 X 32 BLIT SHIELD IF DIR = UP
                    If Player(Index).Shield > 0 Then
                        rec.top = Item(Player(Index).Shield).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 32 BLIT WEAPON IF DIR = UP
                    If Player(Index).Weapon > 0 Then
                        rec.top = Item(Player(Index).Weapon).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 32 BLIT NECKLACE IF DIR = UP
                    If Player(Index).Necklace > 0 Then
                        rec.top = Item(Player(Index).Necklace).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

                ' OTHER 32 X 32 BLIT SPRITE
                If 0 + CustomPlayers = 0 Then
                    rec.top = GetPlayerSprite(Index) * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 32 BLIT ARMOR
                If Player(Index).Armor > 0 Then
                    rec.top = Item(Player(Index).Armor).Pic * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 32 BLIT HELMET
                If Player(Index).Helmet > 0 Then
                    rec.top = Item(Player(Index).Helmet).Pic * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 32 BLIT LEGS
                If Player(Index).legs > 0 Then
                    rec.top = Item(Player(Index).legs).Pic * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                ' OTHER 32 X 32 IF DIR <> UP
                If GetPlayerDir(Index) <> DIR_UP Then

                    ' OTHER 32 X 32 BLIT SHIELD IF DIR <> UP
                    If Player(Index).Shield > 0 Then
                        rec.top = Item(Player(Index).Shield).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 32 BLIT WEAPON IF DIR <> UP
                    If Player(Index).Weapon > 0 Then
                        rec.top = Item(Player(Index).Weapon).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                    ' OTHER 32 X 32 BLIT NECKLACE IF DIR <> UP
                    If Player(Index).Necklace > 0 Then
                        rec.top = Item(Player(Index).Necklace).Pic * PIC_Y
                        rec.Bottom = rec.top + PIC_Y
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If

                End If

            ' OTHER 32 X 32 NON PAPERDOLL
            Else

                ' OTHER 32 X 32 BLIT NON-PD SPRITE
                If 0 + CustomPlayers = 0 Then
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.top = Player(Index).head * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).body * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_body, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    rec.top = Player(Index).leg * PIC_Y
                    rec.Bottom = rec.top + PIC_Y
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_legs, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
            End If

        End If
    End If

End Sub
Sub BltPlayerTop(ByVal Index As Long)
    Dim Anim As Byte
    Dim X As Long, Y As Long, yMod As Long
    Dim AttackSpeed As Long

    If SpriteSize = 1 Then
        If GetPlayerWeaponSlot(Index) > 0 Then
            AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
        Else
            AttackSpeed = 1000
        End If
        
        If WalkFix = 0 Then
            ' Check for animation
            Anim = 0
            If Player(Index).Attacking = 0 Then
                Select Case GetPlayerDir(Index)
                    Case DIR_UP
                        If (Player(Index).yOffset < PIC_Y / 2) Then
                            Anim = 1
                        End If
                    Case DIR_DOWN
                        If (Player(Index).yOffset < PIC_Y / 2 * -1) Then
                            Anim = 1
                        End If
                    Case DIR_LEFT
                        If (Player(Index).xOffset < PIC_Y / 2) Then
                            Anim = 1
                        End If
                    Case DIR_RIGHT
                        If (Player(Index).xOffset < PIC_Y / 2 * -1) Then
                            Anim = 1
                        End If
                End Select
            Else
                If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
                    Anim = 2
                End If
            End If
        Else
            ' Check for animation
              Anim = 1
            If Player(Index).Attacking = 0 Then
              Select Case GetPlayerDir(Index)
                Case DIR_UP
                  If (Player(Index).yOffset > 8) Then Anim = Player(Index).Step
                Case DIR_DOWN
                  If (Player(Index).yOffset < -8) Then Anim = Player(Index).Step
                Case DIR_LEFT
                  If (Player(Index).xOffset > 8) Then Anim = Player(Index).Step
                Case DIR_RIGHT
                  If (Player(Index).xOffset < -8) Then Anim = Player(Index).Step
              End Select
            Else
              If Player(Index).AttackTimer + 1000 > GetTickCount Then
                Anim = 2
              End If
            End If
        End If

        ' Check to see if we want to stop making him attack
        If Player(Index).AttackTimer + AttackSpeed < GetTickCount Then
            Player(Index).Attacking = 0
            'Player(Index).AttackTimer = 0
        End If

        If paperdoll = 1 And GetPlayerPaperdoll(Index) = 1 Then
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
            rec.Right = rec.Left + PIC_X

            If Index = MyIndex Then
                X = NewX + sx
                Y = NewY + sx - 32
                
                ' Fixing "Player head disspear" bug - Emblem
                ' It was caused by trying to blt to a invalid location.
                If Y < 0 Then
                    yMod = Y
                    Y = 0
                End If

                If 0 + CustomPlayers = 0 Then
                    rec.top = GetPlayerSprite(Index) * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.top = GetPlayerHead(Index) * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If GetPlayerDir(Index) = DIR_UP Then
                    If Player(MyIndex).Shield > 0 Then
                        rec.top = Item(Player(MyIndex).Shield).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Weapon > 0 Then
                        rec.top = Item(Player(MyIndex).Weapon).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Necklace > 0 Then
                        rec.top = Item(Player(MyIndex).Necklace).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If

                If Player(MyIndex).Armor > 0 Then
                    rec.top = Item(Player(MyIndex).Armor).Pic * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(MyIndex).legs > 0 Then
                    rec.top = Item(Player(MyIndex).legs).Pic * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(MyIndex).Helmet > 0 Then
                    rec.top = Item(Player(MyIndex).Helmet).Pic * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                If GetPlayerDir(Index) <> DIR_UP Then
                    If Player(MyIndex).Shield > 0 Then
                        rec.top = Item(Player(MyIndex).Shield).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Necklace > 0 Then
                        rec.top = Item(Player(MyIndex).Necklace).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(MyIndex).Weapon > 0 Then
                        rec.top = Item(Player(MyIndex).Weapon).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X, Y, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If


            Else
                X = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset
                Y = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - 32

                If Y < 0 Then
                    yMod = Y
                    Y = 0
                End If

                If 0 + CustomPlayers = 0 Then
                    rec.top = GetPlayerSprite(Index) * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                Else
                    rec.top = GetPlayerHead(Index) * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod

                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If GetPlayerDir(Index) = DIR_UP Then
                    If Player(Index).Shield > 0 Then
                        rec.top = Item(Player(Index).Shield).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Necklace > 0 Then
                        rec.top = Item(Player(Index).Necklace).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Weapon > 0 Then
                        rec.top = Item(Player(Index).Weapon).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If

                If Player(Index).Armor > 0 Then
                    rec.top = Item(Player(Index).Armor).Pic * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(Index).legs > 0 Then
                    rec.top = Item(Player(Index).legs).Pic * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If

                If Player(Index).Helmet > 0 Then
                    rec.top = Item(Player(Index).Helmet).Pic * 64 - yMod
                    rec.Bottom = rec.top + PIC_Y + yMod
                    Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                End If
                If GetPlayerDir(Index) <> DIR_UP Then
                    If Player(Index).Shield > 0 Then
                        rec.top = Item(Player(Index).Shield).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Necklace > 0 Then
                        rec.top = Item(Player(Index).Necklace).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                    If Player(Index).Weapon > 0 Then
                        rec.top = Item(Player(Index).Weapon).Pic * 64 - yMod
                        rec.Bottom = rec.top + PIC_Y + yMod
                        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
                    End If
                End If
            End If
        Else
            If Index = MyIndex Then
                X = NewX + sx
                Y = NewY + sx - 32
                
            Else
                X = (GetPlayerX(Index) - NewPlayerX) * PIC_X + sx + Player(Index).xOffset - NewXOffset
                Y = (GetPlayerY(Index) - NewPlayerY) * PIC_Y + sx + Player(Index).yOffset - NewYOffset - 32
            End If
            
            If Y < 0 Then
                yMod = Y
                Y = 0
            End If
            
            rec.top = GetPlayerSprite(Index) * 64 - yMod
            rec.Bottom = rec.top + PIC_Y + yMod
            rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
            rec.Right = rec.Left + PIC_X
            
            If 0 + CustomPlayers = 0 Then
                Call DD_BackBuffer.BltFast(X, Y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            Else
                rec.top = Player(Index).head * 64 - yMod
                rec.Bottom = rec.top + PIC_Y + yMod
                Call DD_BackBuffer.BltFast(X, Y, DD_player_head, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
            End If
        End If
    Else

        If GetPlayerWeaponSlot(Index) > 0 Then
            AttackSpeed = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AttackSpeed
        Else
            AttackSpeed = 1000
        End If

        ' Check for animation
        Anim = 0
        If Player(Index).Attacking = 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (Player(Index).yOffset < PIC_Y / 2) Then
                        Anim = 1
                    End If
                Case DIR_DOWN
                    If (Player(Index).yOffset < PIC_Y / 2 * -1) Then
                        Anim = 1
                    End If
                Case DIR_LEFT
                    If (Player(Index).xOffset < PIC_Y / 2) Then
                        Anim = 1
                    End If
                Case DIR_RIGHT
                    If (Player(Index).xOffset < PIC_Y / 2 * -1) Then
                        Anim = 1
                    End If
            End Select
        Else
            If Player(Index).AttackTimer + Int(AttackSpeed / 2) > GetTickCount Then
                Anim = 2
            End If
        End If

    End If
End Sub

Sub BltMapNPCName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long

    If Npc(MapNpc(Index).Num).Big = 0 And Npc(MapNpc(Index).Num).SpriteSize = 0 Then
        TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset + CLng(PIC_X / 2) - ((Len(Trim$(Npc(MapNpc(Index).Num).name)) / 2) * 8)
        TextY = MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).yOffset - CLng(PIC_Y / 2) - 4
    Else
        TextX = MapNpc(Index).X * PIC_X + sx + MapNpc(Index).xOffset + CLng(PIC_X / 2) - ((Len(Trim$(Npc(MapNpc(Index).Num).name)) / 2) * 8)
        TextY = MapNpc(Index).Y * PIC_Y + sx + MapNpc(Index).yOffset - CLng(PIC_Y / 2) - 32
    End If

    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, Trim$(Npc(MapNpc(Index).Num).name), vbWhite)
End Sub

Sub BltNpcBody(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim Y As Long
    Dim modify As Long

' Only used if ever want to switch to blt rather then bltfast
' With rec_pos
' .Top = MapNpc(MapNpcNum).Y * PIC_Y + MapNpc(MapNpcNum).YOffset
' .Bottom = .Top + PIC_Y
' .Left = MapNpc(MapNpcNum).X * PIC_X + MapNpc(MapNpcNum).XOffset
' .Right = .Left + PIC_X
' End With

    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
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

    If Npc(MapNpc(MapNpcNum).Num).Big = 1 Then
        rec.top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
        rec.Bottom = rec.top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64

        X = MapNpc(MapNpcNum).X * 32 + sx - 16 + MapNpc(MapNpcNum).xOffset
        Y = MapNpc(MapNpcNum).Y * 32 + sx + MapNpc(MapNpcNum).yOffset

        If Y < 0 Then
            modify = -Y
            rec.top = rec.top + modify
            rec.Bottom = rec.top + 32
            Y = 0
        End If

        If X < 0 Then
            ' rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
            ' modify = -X
            ' rec.Left = rec.Left + modify - 16
            ' rec.Right = rec.Left + 48
            ' X = 0
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

        If Npc(MapNpc(MapNpcNum).Num).SpriteSize = 1 Then
            rec.top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            rec.Right = rec.Left + PIC_X

            X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
            Y = MapNpc(MapNpcNum).Y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset

' Check if its out of bounds because of the offset

            If Y < 0 Then
                rec.top = rec.top + (Y * -1)
                Y = 0
            End If

            ' Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            rec.top = Npc(MapNpc(MapNpcNum).Num).Sprite * PIC_Y
            rec.Bottom = rec.top + PIC_Y
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
            rec.Right = rec.Left + PIC_X

            X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
            Y = MapNpc(MapNpcNum).Y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset

            ' Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim X As Long
    Dim Y As Long
    Dim NPC_number As Long
    Dim modify As Long

    ' Get the NPC number
    NPC_number = MapNpc(MapNpcNum).Num

    If Npc(NPC_number).Big = 0 Then
        If Npc(MapNpc(MapNpcNum).Num).SpriteSize = 0 Then
            Exit Sub
        End If
    End If

' Only used if ever want to switch to blt rather then bltfast
' With rec_pos
' .Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset
' .Bottom = .Top + PIC_Y
' .Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset
' .Right = .Left + PIC_X
' End With

    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).yOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2) Then
                    Anim = 1
                End If
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).xOffset < PIC_Y / 2 * -1) Then
                    Anim = 1
                End If
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
        rec.top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64
        rec.Bottom = rec.top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X

        X = MapNpc(MapNpcNum).X * PIC_X + sx + MapNpc(MapNpcNum).xOffset
        Y = MapNpc(MapNpcNum).Y * PIC_Y + sx + MapNpc(MapNpcNum).yOffset - 32

        ' Check if its out of bounds because of the offset
        If Y < 0 Then
            rec.top = rec.top + (Y * -1)
            Y = 0
        End If

        ' Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(X - (NewPlayerX * PIC_X) - NewXOffset, Y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.top = Npc(MapNpc(MapNpcNum).Num).Sprite * PIC_Y

        rec.top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64
        rec.Bottom = rec.top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64

        X = MapNpc(MapNpcNum).X * 32 + sx - 16 + MapNpc(MapNpcNum).xOffset
        Y = MapNpc(MapNpcNum).Y * 32 + sx - 32 + MapNpc(MapNpcNum).yOffset

        If Y < 0 Then
            modify = -Y
            rec.top = rec.top + modify
            rec.Bottom = rec.top + 32
            Y = 0
        End If

        If X < 0 Then
            ' rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
            ' modify = -X
            ' rec.Left = rec.Left + modify - 16
            ' rec.Right = rec.Left + 48
            ' X = 0
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
            color = QBColor(Val(Player(Index).color - 1))
        End If
    Else
        ' Check access level
        If GetPlayerPK(Index) = NO Then
            color = QBColor(YELLOW)
            Select Case GetPlayerAccess(Index)
                Case 0
                    color = QBColor(BROWN)
                Case 1
                    color = QBColor(DARKGREY)
                Case 2
                    color = QBColor(CYAN)
                Case 3
                    color = QBColor(BLUE)
                Case 4
                    color = QBColor(PINK)
            End Select
        Else
            color = QBColor(BRIGHTRED)
        End If
    End If

    If SpriteSize = 1 Then
        If Index = MyIndex Then
            If lvl >= 1 Then
                TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
            Else
                TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
            End If

            TextY = NewY + sx - 50
            If lvl >= 1 Then
                Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index), color)
            Else
                Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), color)
            End If
        Else
            ' Draw name
            If lvl >= 1 Then
                TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
            Else
                TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
            End If

            TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 32

            If lvl >= 1 Then
                Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index), color)
            Else
                Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), color)
            End If
        End If
    Else
        If SpriteSize = 2 Then
            If Index = MyIndex Then
                If lvl >= 1 Then
                    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
                Else
                    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
                End If

                TextY = NewY + sx - 50
                If lvl >= 1 Then
                    Call DrawText(TexthDC, TextX, TextY - PIC_Y, GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index), color)
                Else
                    Call DrawText(TexthDC, TextX, TextY - PIC_Y, GetPlayerName(MyIndex), color)
                End If
            Else
                ' Draw name
                If lvl >= 1 Then
                    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
                Else
                    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
                End If

                TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 32

                If lvl >= 1 Then
                    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index), color)
                Else
                    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset - PIC_Y, GetPlayerName(Index), color)
                End If
            End If
        Else
            If Index = MyIndex Then
                If lvl >= 1 Then
                    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
                Else
                    TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * 8)
                End If
                TextY = NewY + sx - Int(PIC_Y / 2)

                If lvl >= 1 Then
                    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex) & " lvl: " & GetPlayerLevel(Index), color)
                Else
                    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), color)
                End If
            Else
                ' Draw name
                If lvl >= 1 Then
                    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index) & " lvl: " & GetPlayerLevel(Index)) / 2) * 8)
                Else
                    TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
                End If

                TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2)

                If lvl >= 1 Then
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

    ' Check access level.
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerGuildAccess(Index)
            Case 0
                color = QBColor(RED)
            Case 1
                color = QBColor(BRIGHTCYAN)
            Case 2
                color = QBColor(PINK)
            Case 3
                color = QBColor(BRIGHTGREEN)
            Case 4
                color = QBColor(YELLOW)
        End Select
    Else
        color = QBColor(BRIGHTRED)
    End If

    ' Draw the players guild.
    If Index = MyIndex Then
        TextX = NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerGuild(MyIndex)) / 2) * 8)

        If SpriteSize = 1 Then
            TextY = NewY + sx - Int(PIC_Y / 4) - 52
        Else
            TextY = NewY + sx - Int(PIC_Y / 4) - 20
        End If

        Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), color)
    Else
        TextX = GetPlayerX(Index) * PIC_X + sx + Player(Index).xOffset + Int(PIC_X / 2) - ((Len(GetPlayerGuild(Index)) / 2) * 8)

        If SpriteSize = 1 Then
            TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 44
        Else
            TextY = GetPlayerY(Index) * PIC_Y + sx + Player(Index).yOffset - Int(PIC_Y / 2) - 12
        End If

        Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerGuild(Index), color)
    End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
    If WalkFix = 0 Then
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
                    ' Removed until the server supports SP. [Mellowz]
                    ' Call SetPlayerSP(Index, GetPlayerSP(Index) - 1)
                    frmStable.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
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
                    ' Call AddText("You are to tired to run.", Blue)
                    Player(Index).Moving = MOVING_WALKING
                End If
    
    
                ' Check if completed walking over to the next tile
                If (Player(Index).xOffset = 0) And (Player(Index).yOffset = 0) Then
                    Player(Index).Moving = 0
                End If
            End If
        End If
    'Walkfix starts here
    Else
         ' Check if player is walking, and if so process moving them over
          If Player(Index).Moving = MOVING_WALKING Then
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
         
            ' Check if completed walking over to the next tile
            If Player(Index).Dir = DIR_RIGHT Or Player(Index).Dir = DIR_DOWN Then
              If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                  Player(Index).Step = 2
                Else
                  Player(Index).Step = 0
                End If
              End If
            Else
              If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                  Player(Index).Step = 2
                Else
                  Player(Index).Step = 0
                End If
              End If
            End If
          End If
        
          ' Check if player is running, and if so process moving them over
          If Player(Index).Moving = MOVING_RUNNING Then
            If GetPlayerSP(Index) > 0 Then
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
            Else
                ' Call AddText("You are to tired to run.", Blue)
                Player(Index).Moving = MOVING_WALKING
            End If
         
            ' Check if completed walking over to the next tile
            If Player(Index).Dir = DIR_RIGHT Or Player(Index).Dir = DIR_DOWN Then
              If (Player(Index).xOffset >= 0) And (Player(Index).yOffset >= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                  Player(Index).Step = 2
                Else
                  Player(Index).Step = 0
                End If
              End If
            Else
              If (Player(Index).xOffset <= 0) And (Player(Index).yOffset <= 0) Then
                Player(Index).Moving = 0
                If Player(Index).Step = 0 Then
                  Player(Index).Step = 2
                Else
                  Player(Index).Step = 0
                End If
              End If
            End If
          End If
         
          Select Case GetPlayerDir(Index)
            Case DIR_UP
              If Player(Index).yOffset <= 0 Then
                Player(Index).yOffset = 0
              End If
            Case DIR_DOWN
              If Player(Index).yOffset >= 0 Then
                Player(Index).yOffset = 0
              End If
            Case DIR_LEFT
              If Player(Index).xOffset <= 0 Then
                Player(Index).xOffset = 0
              End If
            Case DIR_RIGHT
              If Player(Index).xOffset >= 0 Then
                Player(Index).xOffset = 0
              End If
          End Select
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
    Dim name As String
    Dim I As Long

    MyText = frmStable.txtMyTextBox.Text

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
        frmStable.txtMyTextBox.Text = vbNullString
                
        If Player(MyIndex).Y - 1 > -1 Then
            If Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                Call AddText("The Sign Reads:", BLUE)
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), GREY)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), GREY)
                End If
                If Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> vbNullString Then
                    Call AddText(Trim$(Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3), GREY)
                End If
                Exit Sub
            End If
        End If
        
        ' Map message
        If frmStable.MapChat.Text = "Map" Then
            'Check if the user uses other chat then mapchat
            ' Broadcast message
            If Not Mid$(MyText, 1, 1) = "/" Then
                If Mid$(MyText, 1, 1) = "'" Then
                    ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                    If Len(Trim$(ChatText)) > 0 Then
                        Call BroadcastMsg(ChatText)
                    End If
                    MyText = vbNullString
                    Exit Sub
                End If
            End If
    
            ' Emote message
            If Not Mid$(MyText, 1, 1) = "/" Then
                If Mid$(MyText, 1, 1) = "-" Then
                    ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                    If Len(Trim$(ChatText)) > 0 Then
                        Call EmoteMsg(ChatText)
                    End If
                    MyText = vbNullString
                    Exit Sub
                End If
            End If
            
            ' Guild message
            If Not Mid$(MyText, 1, 1) = "/" Then
                If Mid$(MyText, 1, 1) = "@" Then
                    ChatText = MyText
                    If Len(Trim$(ChatText)) > 0 Then
                        Call GuildChat(ChatText)
                    End If
                    MyText = vbNullString
                    Exit Sub
                End If
            End If
            
    
            ' Player message
            If Not Mid$(MyText, 1, 1) = "/" Then
                If Mid$(MyText, 1, 1) = "!" Then
                    ChatText = Mid$(MyText, 2, Len(MyText) - 1)
                    name = vbNullString
        
                    ' Get the desired player from the user text
                    For I = 1 To Len(ChatText)
                        If Mid$(ChatText, I, 1) <> " " Then
                            name = name & Mid$(ChatText, I, 1)
                        Else
                            Exit For
                        End If
                    Next
        
                    ' Make sure they are actually sending something
                    If Len(ChatText) - I > 0 Then
                        ChatText = Mid$(ChatText, I + 1, Len(ChatText) - I)
        
                        ' Send the message to the player
                        Call PlayerMsg(ChatText, name)
                    Else
                        Call AddText("Usage: !playername msghere", AlertColor)
                    End If
                    MyText = vbNullString
                    Exit Sub
                End If
                
                If Len(Trim$(MyText)) > 0 Then
                    Call SayMsg(MyText)
                End If
                MyText = vbNullString
                Exit Sub
            End If
        End If
        
        ' Broadcast message
        If frmStable.MapChat.Text = "Global" Then
            ChatText = MyText
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
        
        ' Guild message
        If frmStable.MapChat.Text = "Guild" Then
            ChatText = MyText
            If Len(Trim$(ChatText)) > 0 Then
                Call GuildChat(ChatText)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' Player message
        If frmStable.MapChat.Text = "Private" Then
            ChatText = MyText
            name = vbNullString

            ' Get the desired player from the user text
            For I = 1 To Len(ChatText)
                If Mid$(ChatText, I, 1) <> " " Then
                    name = name & Mid$(ChatText, I, 1)
                Else
                    Exit For
                End If
            Next I
            

            ' Make sure they are actually sending something
            If Len(ChatText) - I > 0 Then
                ChatText = Mid$(ChatText, I + 1, Len(ChatText) - I)

                ' Send the message to the player
                Call PlayerMsg(ChatText, name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            MyText = vbNullString
            Exit Sub
        End If

        ' // Commands //
        ' Verification User
        If LCase$(Mid$(MyText, 1, 5)) = "/info" Then
            ChatText = Mid$(MyText, 7, Len(MyText) - 5)

            If LenB(ChatText) <> 0 Then
                Call SendData("getstats" & SEP_CHAR & ChatText & END_CHAR)
            Else
                Call AddText("Please enter a player name.", BRIGHTRED)
            End If

            MyText = vbNullString
            Exit Sub
        End If

        ' Whos Online
        If LCase$(Mid$(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            MyText = vbNullString
            Exit Sub
        End If
        
        If LCase$(Mid$(MyText, 1, 6)) = "/where" Then
            Call AddText("Map: " & GetPlayerMap(MyIndex) & "; X: " & GetPlayerX(MyIndex) & "; Y: " & GetPlayerY(MyIndex), GREY)
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
        If LCase$(Mid$(MyText, 1, 4)) = "/inv" Then
            frmStable.picInventory.Visible = True
            MyText = vbNullString
            Exit Sub
        End If

        ' Request stats
        If LCase$(Mid$(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & GetPlayerName(MyIndex) & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Refresh Player
        If LCase$(Mid$(MyText, 1, 8)) = "/refresh" Then
            Call SendData("refresh" & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Decline Chat
        If LCase$(Mid$(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData("dchat" & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        ' Accept Chat
        If LCase$(Mid$(MyText, 1, 5)) = "/chat" Then
            Call SendData("achat" & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

        If LCase$(Mid$(MyText, 1, 6)) = "/trade" Then
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
        If LCase$(Mid$(MyText, 1, 7)) = "/accept" Then
            Call SendAcceptTrade
            MyText = vbNullString
            Exit Sub
        End If

        ' Decline Trade
        If LCase$(Mid$(MyText, 1, 8)) = "/decline" Then
            Call SendDeclineTrade
            MyText = vbNullString
            Exit Sub
        End If

        ' Party request
        If LCase$(Mid$(MyText, 1, 6)) = "/party" Then
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
        If LCase$(Mid$(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            MyText = vbNullString
            Exit Sub
        End If

        ' Leave party
        If LCase$(Mid$(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            MyText = vbNullString
            Exit Sub
        End If

        ' House Editor
        If LCase$(Mid$(MyText, 1, 12)) = "/houseeditor" Then
            Call SendRequestEditHouse
            MyText = vbNullString
            Exit Sub
        End If

        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
            ' weather command
            If LCase$(Mid$(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = Mid$(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData("weather" & SEP_CHAR & Val(MyText) & END_CHAR)
                    Else
                        If Trim$(LCase$(MyText)) = "none" Then
                            I = 0
                        End If
                        If Trim$(LCase$(MyText)) = "rain" Then
                            I = 1
                        End If
                        If Trim$(LCase$(MyText)) = "snow" Then
                            I = 2
                        End If
                        If Trim$(LCase$(MyText)) = "thunder" Then
                            I = 3
                        End If
                        Call SendData("weather" & SEP_CHAR & I & END_CHAR)
                    End If
                End If
                MyText = vbNullString
                Exit Sub
            End If


            ' Clearing a house owner
            If LCase$(Mid$(MyText, 1, 11)) = "/clearowner" Then
                Call SendData("clearowner" & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If

            ' Kicking a player
            If LCase$(Mid$(MyText, 1, 5)) = "/kick" Then
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
            If Mid$(MyText, 1, 4) = "/loc" Then
                If BLoc = False Then
                    BLoc = True
                Else
                    BLoc = False
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Map Editor
            If LCase$(Mid$(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                MyText = vbNullString
                Exit Sub
            End If

            ' Map report
            If LCase$(Mid$(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & END_CHAR)
                MyText = vbNullString
                Exit Sub
            End If

            ' Setting sprite
            If LCase$(Mid$(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid$(MyText, 12, Len(MyText) - 11)

                    Call SendSetPlayerSprite(GetPlayerName(MyIndex), Val(MyText))
                End If
                MyText = vbNullString
                Exit Sub
            End If

            ' Setting player sprite
            If LCase$(Mid$(MyText, 1, 16)) = "/setplayersprite" Then
                If Len(MyText) > 19 Then
                    I = Val(Mid$(MyText, 17, 1))

                    MyText = Mid$(MyText, 18, Len(MyText) - 17)
                    Call SendSetPlayerSprite(I, Val(MyText))
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
            If LCase$(Mid$(MyText, 1, 7)) = "/reboot" Then
                Call SendData("reboot" & END_CHAR)
                Call GlobalMsg("An Administrator has started a server reboot, please log off!")
                MyText = vbNullString
                Exit Sub
            End If

            ' Banning a player
            If LCase$(Mid$(MyText, 1, 4)) = "/ban" Then
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
            If LCase$(Mid$(MyText, 1, 9)) = "/daynight" Then
                Call SendData("daynight" & END_CHAR)
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

            ' Editing arrow request
            If Mid$(MyText, 1, 13) = "/editarrow" Then
                Call SendRequestEditArrow
                MyText = vbNullString
                Exit Sub
            End If

            ' Editing npc request
            If Mid$(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNPC
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
            If LCase$(Trim$(MyText)) = "/editspell" Then
                Call SendRequestEditSpell
                MyText = vbNullString
                Exit Sub
            End If
        End If

        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            ' Giving another player access
            If LCase$(Mid$(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                I = Val(Mid$(MyText, 12, 1))

                MyText = Mid$(MyText, 14, Len(MyText) - 13)

                Call SendSetAccess(MyText, I)
                MyText = vbNullString
                Exit Sub
            End If
            
            'If MyText = "/editmain" Or MyText = "/maineditor" Then
                'Call SendRequestEditMain(
                'Exit Sub
            'End If

            ' Reload Scripts
            If LCase$(Trim$(MyText)) = "/reload" Then
                Call SendReloadScripts
                MyText = vbNullString
                Exit Sub
            End If

            ' Ban destroy
            If LCase$(Mid$(MyText, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                MyText = vbNullString
                Exit Sub
            End If
        End If

        ' Tell them its not a valid command
        If Left$(Trim$(MyText), 1) = "/" Then
            For I = 0 To MAX_EMOTICONS
                If Trim$(Emoticons(I).Command) = Trim$(MyText) And Trim$(Emoticons(I).Command) <> "/" Then
                    Call SendData("checkemoticons" & SEP_CHAR & I & END_CHAR)
                    MyText = vbNullString
                    Exit Sub
                End If
            Next I
            Call SendData("checkcommands" & SEP_CHAR & MyText & END_CHAR)
            MyText = vbNullString
            Exit Sub
        End If

    End If

    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If Len(MyText) > 0 Then
        ' MyText = mid$(MyText, 1, Len(MyText) - 1)
        End If
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) Then
        ' Make sure they just use standard keys, no gay shitty macro keys
        If KeyAscii >= 32 And KeyAscii <= 255 Then
        ' frmMirage.txtMyTextBox.Text = frmMirage.txtMyTextBox.Text & Chr$(KeyAscii)
        ' MyText = MyText & Chr$(KeyAscii)
        End If
    End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim$(MyText) = vbNullString Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & END_CHAR)
    End If
End Sub

Sub CheckAttack()
    Dim AttackSpeed As Long

    If GetPlayerWeaponSlot(MyIndex) > 0 Then
        AttackSpeed = Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AttackSpeed
    Else
        AttackSpeed = 1000
    End If

    If ControlDown Then
        If Player(MyIndex).AttackTimer + AttackSpeed < GetTickCount Then
            If Player(MyIndex).Attacking = 0 Then
                Player(MyIndex).Attacking = 1
                Player(MyIndex).AttackTimer = GetTickCount
                Call SendData("attack" & END_CHAR)
            End If
        End If
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If Not GettingMap Then
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
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True
    End If
End Function

Function CanMove() As Boolean
    Dim I As Long
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

    If Y < 0 Then
        If Map(GetPlayerMap(MyIndex)).Up > 0 Then
            Call SendPlayerRequestNewMap(DIR_UP)
            GettingMap = True
        End If
        CanMove = False
        Exit Function
    ElseIf Y > MAX_MAPY Then
        If Map(GetPlayerMap(MyIndex)).Down > 0 Then
            Call SendPlayerRequestNewMap(DIR_DOWN)
            GettingMap = True
        End If
        CanMove = False
        Exit Function
    ElseIf X < 0 Then
        If Map(GetPlayerMap(MyIndex)).Left > 0 Then
            Call SendPlayerRequestNewMap(DIR_LEFT)
            GettingMap = True
        End If
        CanMove = False
        Exit Function
    ElseIf X > MAX_MAPX Then
        If Map(GetPlayerMap(MyIndex)).Right > 0 Then
            Call SendPlayerRequestNewMap(DIR_RIGHT)
            GettingMap = True
        End If
        CanMove = False
        Exit Function
    End If

    If Not GetPlayerDir(MyIndex) = LAST_DIR Then
        LAST_DIR = GetPlayerDir(MyIndex)
        Call SendPlayerDir
    End If

    If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_BLOCKED Or Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_SIGN Or Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_ROOFBLOCK Then
        CanMove = False
        Exit Function
    End If

    If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_CBLOCK Then
        If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data1 = Player(MyIndex).Class Then
            Exit Function
        End If
        If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data2 = Player(MyIndex).Class Then
            Exit Function
        End If
        If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Data3 = Player(MyIndex).Class Then
            Exit Function
        End If
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
        For I = 1 To MAX_PLAYERS
            If IsPlaying(I) Then
                If GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                    If GetPlayerX(I) = X Then
                        If GetPlayerY(I) = Y Then
                            CanMove = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next I
    End If

    For I = 1 To MAX_MAP_NPCS
        If MapNpc(I).Num > 0 Then
            If MapNpc(I).X = X Then
                If MapNpc(I).Y = Y Then
                    CanMove = False
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Sub CheckMovement()
    Dim s2kX As Integer, s2kY As Integer   ' used below for temp store of X/Y
    
    If Not GettingMap Then
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
                       Player(MyIndex).yOffset = PIC_Y
                       Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
  
                   Case DIR_DOWN
                       Player(MyIndex).yOffset = PIC_Y * -1
                       Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
  
                   Case DIR_LEFT
                       Player(MyIndex).xOffset = PIC_X
                       Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
  
                   Case DIR_RIGHT
                       Player(MyIndex).xOffset = PIC_X * -1
                       Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
               End Select
               
               Call SendPlayerMove     '090829 moved here
  
               s2kX = GetPlayerX(MyIndex)  '090829
               s2kY = GetPlayerY(MyIndex)  '090829

                ' Gotta check :)
                If Map(GetPlayerMap(MyIndex)).Tile(s2kX, s2kY).Type = TILE_TYPE_WARP Or s2kX < 0 Or s2kX > MAX_MAPX Or s2kY < 0 Or s2kY > MAX_MAPY Then
                    GettingMap = True
                End If
                
            End If
        End If
    End If
End Sub

Function FindPlayer(ByVal name As String) As Long
    Dim I As Long

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(I)) >= Len(Trim$(name)) Then
                If UCase$(Mid$(GetPlayerName(I), 1, Len(Trim$(name)))) = UCase$(Trim$(name)) Then
                    FindPlayer = I
                    Exit Function
                End If
            End If
        End If
    Next I

    FindPlayer = 0
End Function

Public Sub UpdateTradeInventory()
    Dim I As Long

    frmPlayerTrade.PlayerInv1.Clear

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 And GetPlayerInvItemNum(MyIndex, I) <= MAX_ITEMS Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, I)).Stackable = 1 Then
                frmPlayerTrade.PlayerInv1.addItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
            Else
                If GetPlayerWeaponSlot(MyIndex) = I Or GetPlayerArmorSlot(MyIndex) = I Or GetPlayerHelmetSlot(MyIndex) = I Or GetPlayerShieldSlot(MyIndex) = I Or GetPlayerLegsSlot(MyIndex) = I Or GetPlayerRingSlot(MyIndex) = I Or GetPlayerNecklaceSlot(MyIndex) = I Then
                    frmPlayerTrade.PlayerInv1.addItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (worn)"
                Else
                    frmPlayerTrade.PlayerInv1.addItem I & ": " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name)
                End If
            End If
        Else
            frmPlayerTrade.PlayerInv1.addItem "<Nothing>"
        End If
    Next I

    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If CurX >= 0 And CurX <= MAX_MAPX Then
        If CurY >= 0 And CurY <= MAX_MAPY Then
            ' Disabled until we get a better movement system. [Mellowz]
            ' Call MoveCharacter(CurX, CurY)
            Call SendData("search" & SEP_CHAR & CurX & SEP_CHAR & CurY & END_CHAR)
        End If
    End If
End Sub

Sub PlayerSearch2(Button As Integer, Shift As Integer, X As Long, Y As Long)
    If CurX >= 0 And CurX <= MAX_MAPX Then
        If CurY >= 0 And CurY <= MAX_MAPY Then
            Call SendData("search2" & SEP_CHAR & CurX & SEP_CHAR & CurY & END_CHAR)
        End If
    End If
End Sub

Public Sub UpdateVisInv()
    Dim Index As Long
    Dim d As Long

    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) <> Index Then
            frmStable.ShieldImage.Picture = LoadPicture()
        End If
        If GetPlayerWeaponSlot(MyIndex) <> Index Then
            frmStable.WeaponImage.Picture = LoadPicture()
        End If
        If GetPlayerHelmetSlot(MyIndex) <> Index Then
            frmStable.HelmetImage.Picture = LoadPicture()
        End If
        If GetPlayerArmorSlot(MyIndex) <> Index Then
            frmStable.ArmorImage.Picture = LoadPicture()
        End If
        If GetPlayerLegsSlot(MyIndex) <> Index Then
            frmStable.LegsImage.Picture = LoadPicture()
        End If
        If GetPlayerRingSlot(MyIndex) <> Index Then
            frmStable.RingImage.Picture = LoadPicture()
        End If
        If GetPlayerNecklaceSlot(MyIndex) <> Index Then
            frmStable.NecklaceImage.Picture = LoadPicture()
        End If
    Next Index

    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = Index Then
            Call BitBlt(frmStable.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmStable.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        End If
        If GetPlayerWeaponSlot(MyIndex) = Index Then
            Call BitBlt(frmStable.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmStable.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        End If
        If GetPlayerHelmetSlot(MyIndex) = Index Then
            Call BitBlt(frmStable.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmStable.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        End If
        If GetPlayerArmorSlot(MyIndex) = Index Then
            Call BitBlt(frmStable.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmStable.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        End If
        If GetPlayerLegsSlot(MyIndex) = Index Then
            Call BitBlt(frmStable.LegsImage.hDC, 0, 0, PIC_X, PIC_Y, frmStable.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        End If
        If GetPlayerRingSlot(MyIndex) = Index Then
            Call BitBlt(frmStable.RingImage.hDC, 0, 0, PIC_X, PIC_Y, frmStable.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        End If
        If GetPlayerNecklaceSlot(MyIndex) = Index Then
            Call BitBlt(frmStable.NecklaceImage.hDC, 0, 0, PIC_X, PIC_Y, frmStable.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        End If
    Next Index

    frmStable.EquipS(0).Visible = False
    frmStable.EquipS(1).Visible = False
    frmStable.EquipS(2).Visible = False
    frmStable.EquipS(3).Visible = False
    frmStable.EquipS(4).Visible = False
    frmStable.EquipS(5).Visible = False
    frmStable.EquipS(6).Visible = False


    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).Num > 0 Then
            If Not Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
                ' frmMirage.descName.Caption = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
                ' Else
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(0).Visible = True
                    frmStable.EquipS(0).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(0).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(1).Visible = True
                    frmStable.EquipS(1).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(1).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(2).Visible = True
                    frmStable.EquipS(2).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(2).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(3).Visible = True
                    frmStable.EquipS(3).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(3).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(4).Visible = True
                    frmStable.EquipS(4).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(4).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(5).Visible = True
                    frmStable.EquipS(5).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(5).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerNecklaceSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(6).Visible = True
                    frmStable.EquipS(6).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(6).Left = frmStable.picInv(d).Left - 2
                Else
                ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
                End If
            End If
        End If
    Next d
End Sub

Public Sub UpdateotherVisInv()
    Dim Index As Long
    Dim d As Long

    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).Num > 0 Then
            If Not Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
                ' frmMirage.descName.Caption = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (" & GetPlayerInvItemValue(MyIndex, d + 1) & ")"
                ' Else
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(0).Visible = True
                    frmStable.EquipS(0).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(0).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(1).Visible = True
                    frmStable.EquipS(1).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(1).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(2).Visible = True
                    frmStable.EquipS(2).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(2).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(3).Visible = True
                    frmStable.EquipS(3).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(3).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerLegsSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(4).Visible = True
                    frmStable.EquipS(4).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(4).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(5).Visible = True
                    frmStable.EquipS(5).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(5).Left = frmStable.picInv(d).Left - 2
                ElseIf GetPlayerNecklaceSlot(MyIndex) = d + 1 Then
                    ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name) & " (worn)"
                    frmStable.EquipS(6).Visible = True
                    frmStable.EquipS(6).top = frmStable.picInv(d).top - 2
                    frmStable.EquipS(6).Left = frmStable.picInv(d).Left - 2
                Else
                ' frmMirage.picInv(d).ToolTipText = trim$(Item(GetPlayerInvItemNum(MyIndex, d + 1)).Name)
                End If
            End If
        End If
    Next d
End Sub


Sub SendGameTime()
    Dim Packet As String

    Packet = "GmTime" & SEP_CHAR & GameTime & END_CHAR
    Call SendData(Packet)
End Sub

Sub UpdateBank()
    Dim I As Long

    frmBank.lstInventory.Clear
    frmBank.lstBank.Clear

    For I = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, I) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, I)).Stackable = 1 Then
                frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name) & " (" & GetPlayerInvItemValue(MyIndex, I) & ")"
            Else
                frmBank.lstInventory.addItem I & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, I)).name)
            End If
        Else
            frmBank.lstInventory.addItem I & "> Empty"
        End If
    Next I

    For I = 1 To MAX_BANK
        If GetPlayerBankItemNum(MyIndex, I) > 0 Then
            If Item(GetPlayerBankItemNum(MyIndex, I)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, I)).Stackable = 1 Then
                frmBank.lstBank.addItem I & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, I)).name) & " (" & GetPlayerBankItemValue(MyIndex, I) & ")"
            Else
                frmBank.lstBank.addItem I & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, I)).name)
            End If
        Else
            frmBank.lstBank.addItem I & "> Empty"
        End If
    Next I

    frmBank.lstBank.ListIndex = 0
    frmBank.lstInventory.ListIndex = 0
End Sub

Sub UseItem()
    Dim d As Long

    Call SendUseItem(Inventory)

    For d = 1 To MAX_INV
        If Player(MyIndex).Inv(d).Num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d)).Type = ITEM_TYPE_POTIONADDMP Or ITEM_TYPE_POTIONADDHP Or ITEM_TYPE_POTIONADDSP Or ITEM_TYPE_POTIONSUBHP Or ITEM_TYPE_POTIONSUBMP Or ITEM_TYPE_POTIONSUBSP Then
                frmStable.picInv(d - 1).Picture = LoadPicture()
            End If
        End If
    Next d

    Call UpdateVisInv
End Sub

Sub DropItem()
    Dim InvNum As Long
    Dim GoldAmount As String

    On Error GoTo DropItem_Error

    InvNum = Inventory

    If GetPlayerInvItemNum(MyIndex, InvNum) > 0 And GetPlayerInvItemNum(MyIndex, InvNum) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Bound = 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, InvNum)).Stackable = 1 Then
                GoldAmount = InputBox("How much " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name) & "(" & GetPlayerInvItemValue(MyIndex, InvNum) & ") would you like to drop?", "Drop " & Trim$(Item(GetPlayerInvItemNum(MyIndex, InvNum)).name), 0, frmStable.Left, frmStable.top)

                If IsNumeric(GoldAmount) Then
                    Call SendDropItem(InvNum, GoldAmount)
                End If
            Else
                Call SendDropItem(InvNum, 0)
            End If
        End If
    End If

    frmStable.picInv(InvNum - 1).Picture = LoadPicture()

    Call UpdateVisInv

    Exit Sub

DropItem_Error:
    Call AddText("Please enter a valid amount for that item!", BRIGHTRED)
End Sub

' Sets the speed of a character based on speed
Sub SetSpeed(ByVal run As String, ByVal speed As Long)
    If LCase$(run) = "walk" Then
        SS_WALK_SPEED = speed
    ElseIf LCase$(run) = "run" Then
        SS_RUN_SPEED = speed
    End If
' Ignore all other cases
End Sub

Sub MoveCharacter(ByVal MX As Integer, ByVal MY As Integer)
    If Player(MyIndex).input = 0 Then
        Exit Sub
    End If
    If GetPlayerY(MyIndex) = MAX_MAPY Then
        If MY = GetPlayerY(MyIndex) Then
            Call SendChangeDir
        End If
    Else
        If MY > GetPlayerY(MyIndex) And Val(MY - GetPlayerY(MyIndex)) > Val(MX - GetPlayerX(MyIndex)) Then
            Call SetPlayerDir(MyIndex, 1)
            If CanMove = True Then
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                DirDown = True
                Call SendPlayerMoveMouse
                Exit Sub
            End If
        End If
    End If

    If GetPlayerY(MyIndex) = 0 Then
        If MY = GetPlayerY(MyIndex) Then
            Call SendChangeDir
        End If
    Else
        If MY < GetPlayerY(MyIndex) And Val(MY - GetPlayerY(MyIndex)) < Val(MX - GetPlayerX(MyIndex)) Then
            Call SetPlayerDir(MyIndex, 0)
            If CanMove = True Then
                Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                DirUp = True
                Call SendPlayerMoveMouse
                Exit Sub
            End If
        End If
    End If

    If GetPlayerX(MyIndex) + 1 = MAX_MAPX Then
        If MX = GetPlayerX(MyIndex) Then
            Call SendChangeDir
        End If
    Else
        If MX > GetPlayerX(MyIndex) And Val(MY - GetPlayerY(MyIndex)) < Val(MX - GetPlayerX(MyIndex)) Then
            Call SetPlayerDir(MyIndex, 3)
            If CanMove = True Then
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                DirRight = True
                Call SendPlayerMoveMouse
                Exit Sub
            End If
        End If

    End If

    If GetPlayerX(MyIndex) = 0 Then
        If MX = GetPlayerX(MyIndex) Then
            Call SendChangeDir
        End If
    Else
        If MX < GetPlayerX(MyIndex) And Val(MY - GetPlayerY(MyIndex)) > Val(MX - GetPlayerX(MyIndex)) Then
            Call SetPlayerDir(MyIndex, 2)
            If CanMove = True Then
                Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                DirLeft = True
                Call SendPlayerMoveMouse
                Exit Sub
            End If
        End If
    End If
End Sub

Public Sub AlwaysOnTop(FormName As Form, bOnTop As Boolean)
    If Not bOnTop Then
        Call SetWindowPos(FormName.hWnd, HWND_TOPMOST, 0, 0, 0, 0, flags)
    Else
        Call SetWindowPos(FormName.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, flags)
    End If
End Sub

Sub GoShop(ByVal Shop As Integer)
    ' Close any other shop windows
    frmNewShop.Hide

    ' Initialize the shop
    Call frmNewShop.loadShop(Shop)
    snumber = Shop

    ' Hide panel
    frmNewShop.picItemInfo.Visible = False

    ' Show shop
    frmNewShop.Show vbModeless, frmStable


    On Error Resume Next
    
    ' Set focus
    frmNewShop.SetFocus

    ' Show page 1 (it starts from 0)
    frmNewShop.showPage (0)
End Sub

Sub IncrementGameClock()
    Dim CurTime As String

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
        CurTime = CStr(Hours - 12)
    Else
        CurTime = Hours
    End If

    If Minutes < 10 Then
        CurTime = CurTime & ":0" & Minutes
    Else
        CurTime = CurTime & ":" & Minutes
    End If

    If Seconds < 10 Then
        CurTime = CurTime & ":0" & Seconds
    Else
        CurTime = CurTime & ":" & Seconds
    End If

    If Hours > 12 Then
        CurTime = CurTime & " PM"
    Else
        CurTime = CurTime & " AM"
    End If

    frmStable.lblGameClock.Caption = CurTime
End Sub

' Returns true if the tile is a roof tile and the player is under that section of roof
Function IsTileRoof(ByVal X As Integer, ByVal Y As Integer) As Boolean
    Dim IsRoof As Boolean
    
    If Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_ROOF Or Map(GetPlayerMap(MyIndex)).Tile(X, Y).Type = TILE_TYPE_ROOFBLOCK Then 'If the tile is a roof or a roofblock
        If Map(GetPlayerMap(MyIndex)).Tile(X, Y).String1 = Map(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).String1 Then 'If the roof ID is the same
            IsTileRoof = True
            Exit Function
        End If
    End If

    IsTileRoof = False
End Function

