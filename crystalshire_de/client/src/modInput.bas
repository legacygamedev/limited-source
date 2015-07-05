Attribute VB_Name = "modInput"
Option Explicit
' keyboard input
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Sub HandleMouseMove(ByVal x As Long, ByVal y As Long, ByVal Button As Long)
Dim i As Long

    ' Set the global cursor position
    GlobalX = x
    GlobalY = y
    GlobalX_Map = GlobalX + (TileView.left * PIC_X) + Camera.left
    GlobalY_Map = GlobalY + (TileView.top * PIC_Y) + Camera.top
    
    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For i = 1 To GUI_Count
            If (x >= GUIWindow(i).x And x <= GUIWindow(i).x + GUIWindow(i).Width) And (y >= GUIWindow(i).y And y <= GUIWindow(i).y + GUIWindow(i).height) Then
                If GUIWindow(i).visible Then
                    Select Case i
                        Case GUI_CHAT, GUI_BARS
                            ' Put nothing here and we can click through them!
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
    
    ' Handle the events
    CurX = TileView.left + ((x + Camera.left) \ PIC_X)
    CurY = TileView.top + ((y + Camera.top) \ PIC_Y)

    If InMapEditor Then
        If Button = vbLeftButton Or Button = vbRightButton Then
            Call MapEditorMouseDown(Button, x, y)
        End If
    End If
End Sub

Public Sub HandleMouseDown(ByVal Button As Long)
Dim i As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For i = 1 To GUI_Count
            If (GlobalX >= GUIWindow(i).x And GlobalX <= GUIWindow(i).x + GUIWindow(i).Width) And (GlobalY >= GUIWindow(i).y And GlobalY <= GUIWindow(i).y + GUIWindow(i).height) Then
                If GUIWindow(i).visible Then
                    Select Case i
                        Case GUI_INVENTORY
                            Inventory_MouseDown Button
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_MouseDown Button
                            Exit Sub
                        Case GUI_MENU
                            Menu_MouseDown Button
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_MouseDown Button
                            Exit Sub
                        Case GUI_MAINMENU
                            MainMenu_MouseDown Button
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_MouseDown
                            Exit Sub
                        Case GUI_CHAT
                            If inChat Then
                                Chat_MouseDown
                                Exit Sub
                            End If
                            If inTutorial Then
                                Tutorial_MouseDown
                                Exit Sub
                            End If
                        Case GUI_SHOP
                            Shop_MouseDown
                            Exit Sub
                        Case GUI_PARTY
                            Party_MouseDown
                            Exit Sub
                        Case GUI_BARS
                            ' nothing here so we can click through
                        Case GUI_OPTIONS
                            Options_MouseDown
                            Exit Sub
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
        ' check chat buttons
        If Not inChat And Not inTutorial Then
            ChatScroll_MouseDown
        End If
    End If
    
    ' Handle events
    If InMapEditor Then
        Call MapEditorMouseDown(Button, GlobalX, GlobalY, False)
    Else
        ' left click
        If Button = vbLeftButton Then
            ' targetting
            'Call PlayerSearch(CurX, CurY)
            FindTarget
        ' right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= 2 Then AdminWarp CurX, CurY
            End If
        End If
    End If
End Sub

Public Sub HandleMouseUp(ByVal Button As Long)
Dim i As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For i = 1 To GUI_Count
            If (GlobalX >= GUIWindow(i).x And GlobalX <= GUIWindow(i).x + GUIWindow(i).Width) And (GlobalY >= GUIWindow(i).y And GlobalY <= GUIWindow(i).y + GUIWindow(i).height) Then
                If GUIWindow(i).visible Then
                    Select Case i
                        Case GUI_INVENTORY
                            Inventory_MouseUp
                        Case GUI_SPELLS
                            Spells_MouseUp
                        Case GUI_MENU
                            Menu_MouseUp
                        Case GUI_HOTBAR
                            Hotbar_MouseUp
                        Case GUI_MAINMENU
                            MainMenu_MouseUp
                        Case GUI_CHARACTER
                            Character_MouseUp
                        Case GUI_CHAT
                            If inChat Then
                                Chat_MouseUp
                            End If
                            If inTutorial Then
                                Tutorial_MouseUp
                            End If
                        Case GUI_SHOP
                            Shop_MouseUp
                        Case GUI_PARTY
                            Party_MouseUp
                        Case GUI_OPTIONS
                            Options_MouseUp
                    End Select
                End If
            End If
        Next
    End If

    ' Stop dragging if we haven't catched it already
    DragInvSlotNum = 0
    DragBankSlotNum = 0
    DragSpell = 0
    ' reset buttons
    resetClickedButtons
    ' stop scrolling chat
    ChatButtonUp = False
    ChatButtonDown = False
End Sub

Public Sub HandleDoubleClick()
Dim i As Long

    ' GUI processing
    If Not InMapEditor And Not hideGUI Then
        For i = 1 To GUI_Count
            If (GlobalX >= GUIWindow(i).x And GlobalX <= GUIWindow(i).x + GUIWindow(i).Width) And (GlobalY >= GUIWindow(i).y And GlobalY <= GUIWindow(i).y + GUIWindow(i).height) Then
                If GUIWindow(i).visible Then
                    Select Case i
                        Case GUI_INVENTORY
                            Inventory_DoubleClick
                            Exit Sub
                        Case GUI_SPELLS
                            Spells_DoubleClick
                            Exit Sub
                        Case GUI_CHARACTER
                            Character_DoubleClick
                            Exit Sub
                        Case GUI_HOTBAR
                            Hotbar_DoubleClick
                        Case GUI_SHOP
                            Shop_DoubleClick
                        Case Else
                            Exit Sub
                    End Select
                End If
            End If
        Next
    End If
End Sub

Public Sub OpenGuiWindow(ByVal index As Long)
    If index = 1 Then
        GUIWindow(GUI_INVENTORY).visible = Not GUIWindow(GUI_INVENTORY).visible
    Else
        GUIWindow(GUI_INVENTORY).visible = False
    End If
    
    If index = 2 Then
        GUIWindow(GUI_SPELLS).visible = Not GUIWindow(GUI_SPELLS).visible
    Else
        GUIWindow(GUI_SPELLS).visible = False
    End If
    
    If index = 3 Then
        GUIWindow(GUI_CHARACTER).visible = Not GUIWindow(GUI_CHARACTER).visible
    Else
        GUIWindow(GUI_CHARACTER).visible = False
    End If
    
    If index = 4 Then
        GUIWindow(GUI_OPTIONS).visible = Not GUIWindow(GUI_OPTIONS).visible
    Else
        GUIWindow(GUI_OPTIONS).visible = False
    End If
    
    If index = 6 Then
        GUIWindow(GUI_PARTY).visible = Not GUIWindow(GUI_PARTY).visible
    Else
        GUIWindow(GUI_PARTY).visible = False
    End If
End Sub

' Tutorial
Public Sub Tutorial_MouseDown()
Dim i As Long, x As Long, y As Long, Width As Long
    
    For i = 1 To 4
        If Len(Trim$(chatOpt(i))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(i)) & "]")
            x = GUIWindow(GUI_CHAT).x + 200 + (130 - (Width / 2))
            y = GUIWindow(GUI_CHAT).y + 115 - ((i - 1) * 15)
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                chatOptState(i) = 2 ' clicked
            End If
        End If
    Next
End Sub

Public Sub Tutorial_MouseUp()
Dim i As Long, x As Long, y As Long, Width As Long
    
    For i = 1 To 4
        If Len(Trim$(chatOpt(i))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(i)) & "]")
            x = GUIWindow(GUI_CHAT).x + 200 + (130 - (Width / 2))
            y = GUIWindow(GUI_CHAT).y + 115 - ((i - 1) * 15)
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                ' are we clicked?
                If chatOptState(i) = 2 Then
                    SetTutorialState tutorialState + 1
                    ' play sound
                    Play_Sound Sound_ButtonClick
                End If
            End If
        End If
    Next
    
    For i = 1 To 4
        chatOptState(i) = 0 ' normal
    Next
End Sub

' Npc Chat
Public Sub Chat_MouseDown()
Dim i As Long, x As Long, y As Long, Width As Long
    
    For i = 1 To 4
        If Len(Trim$(chatOpt(i))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(i)) & "]")
            x = GUIWindow(GUI_CHAT).x + 95 + (155 - (Width / 2))
            y = GUIWindow(GUI_CHAT).y + 115 - ((i - 1) * 15)
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                chatOptState(i) = 2 ' clicked
            End If
        End If
    Next
End Sub

Public Sub Chat_MouseUp()
Dim i As Long, x As Long, y As Long, Width As Long
    
    For i = 1 To 4
        If Len(Trim$(chatOpt(i))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(i)) & "]")
            x = GUIWindow(GUI_CHAT).x + 95 + (155 - (Width / 2))
            y = GUIWindow(GUI_CHAT).y + 115 - ((i - 1) * 15)
            If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                ' are we clicked?
                If chatOptState(i) = 2 Then
                    SendChatOption i
                    ' play sound
                    Play_Sound Sound_ButtonClick
                End If
            End If
        End If
    Next
    
    For i = 1 To 4
        chatOptState(i) = 0 ' normal
    Next
End Sub

' scroll bar
Public Sub ChatScroll_MouseDown()
Dim i As Long, x As Long, y As Long, Width As Long
    
    ' find out which button we're clicking
    For i = 34 To 35
        x = GUIWindow(GUI_CHAT).x + Buttons(i).x
        y = GUIWindow(GUI_CHAT).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            Buttons(i).state = 2 ' clicked
            ' scroll the actual chat
            Select Case i
                Case 34 ' up
                    'ChatScroll = ChatScroll + 1
                    ChatButtonUp = True
                Case 35 ' down
                    'ChatScroll = ChatScroll - 1
                    'If ChatScroll < 8 Then ChatScroll = 8
                    ChatButtonDown = True
            End Select
        End If
    Next
End Sub

' Shop
Public Sub Shop_MouseUp()
Dim i As Long, x As Long, y As Long, Buffer As clsBuffer

    ' find out which button we're clicking
    For i = 23 To 23
        x = GUIWindow(GUI_SHOP).x + Buttons(i).x
        y = GUIWindow(GUI_SHOP).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            If Buttons(i).state = 2 Then
                ' do stuffs
                Select Case i
                    Case 23
                        ' exit
                        Set Buffer = New clsBuffer
                        Buffer.WriteLong CCloseShop
                        SendData Buffer.ToArray()
                        Set Buffer = Nothing
                        GUIWindow(GUI_SHOP).visible = False
                        InShop = 0
                End Select
                ' play sound
                Play_Sound Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Shop_MouseDown()
Dim i As Long, x As Long, y As Long

    ' find out which button we're clicking
    For i = 23 To 23
        x = GUIWindow(GUI_SHOP).x + Buttons(i).x
        y = GUIWindow(GUI_SHOP).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            Buttons(i).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Shop_DoubleClick()
Dim shopSlot As Long

    shopSlot = IsShopItem(GlobalX, GlobalY)

    If shopSlot > 0 Then
        ' buy item code
        BuyItem shopSlot
    End If
End Sub

' Party
Public Sub Party_MouseUp()
Dim i As Long, x As Long, y As Long, Buffer As clsBuffer

    ' find out which button we're clicking
    For i = 24 To 25
        x = GUIWindow(GUI_PARTY).x + Buttons(i).x
        y = GUIWindow(GUI_PARTY).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            If Buttons(i).state = 2 Then
                ' do stuffs
                Select Case i
                    Case 24 ' invite
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendPartyRequest
                        Else
                            AddText "Invalid invitation target.", BrightRed
                        End If
                    Case 25 ' leave
                        If Party.Leader > 0 Then
                            SendPartyLeave
                        Else
                            AddText "You are not in a party.", BrightRed
                        End If
                End Select
                ' play sound
                Play_Sound Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Party_MouseDown()
Dim i As Long, x As Long, y As Long
    ' find out which button we're clicking
    For i = 24 To 25
        x = GUIWindow(GUI_PARTY).x + Buttons(i).x
        y = GUIWindow(GUI_PARTY).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            Buttons(i).state = 2 ' clicked
        End If
    Next
End Sub

'Options
Public Sub Options_MouseUp()
Dim i As Long, x As Long, y As Long, Buffer As clsBuffer, layerNum As Long

    ' find out which button we're clicking
    For i = 26 To 33
        x = GUIWindow(GUI_OPTIONS).x + Buttons(i).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            If Buttons(i).state = 3 Then
                ' do stuffs
                Select Case i
                    Case 26 ' music on
                        Options.Music = 1
                        Play_Music Trim$(Map.Music)
                        SaveOptions
                        Buttons(26).state = 2
                        Buttons(27).state = 0
                    Case 27 ' music off
                        Options.Music = 0
                        Stop_Music
                        SaveOptions
                        Buttons(26).state = 0
                        Buttons(27).state = 2
                    Case 28 ' sound on
                        Options.sound = 1
                        SaveOptions
                        Buttons(28).state = 2
                        Buttons(29).state = 0
                    Case 29 ' sound off
                        Options.sound = 0
                        SaveOptions
                        Buttons(28).state = 0
                        Buttons(29).state = 2
                    Case 30 ' debug on
                        'Options.Debug = 1
                        'SaveOptions
                        'Buttons(30).state = 2
                        'Buttons(31).state = 0
                    Case 31 ' debug off
                        'Options.Debug = 0
                        'SaveOptions
                        'Buttons(30).state = 0
                        'Buttons(31).state = 2
                    Case 32 ' noAuto on
                        Options.noAuto = 0
                        SaveOptions
                        Buttons(32).state = 2
                        Buttons(33).state = 0
                        ' cache render state
                        For x = 0 To Map.MaxX
                            For y = 0 To Map.MaxY
                                For layerNum = 1 To MapLayer.Layer_Count - 1
                                    cacheRenderState x, y, layerNum
                                Next
                            Next
                        Next
                    Case 33 ' noAuto off
                        Options.noAuto = 1
                        SaveOptions
                        Buttons(32).state = 0
                        Buttons(33).state = 2
                        ' cache render state
                        For x = 0 To Map.MaxX
                            For y = 0 To Map.MaxY
                                For layerNum = 1 To MapLayer.Layer_Count - 1
                                    cacheRenderState x, y, layerNum
                                Next
                            Next
                        Next
                End Select
                ' play sound
                Play_Sound Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Options_MouseDown()
Dim i As Long, x As Long, y As Long
    ' find out which button we're clicking
    For i = 26 To 33
        x = GUIWindow(GUI_OPTIONS).x + Buttons(i).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            If Buttons(i).state = 0 Then
                Buttons(i).state = 3 ' clicked
            End If
        End If
    Next
End Sub

' Menu
Public Sub Menu_MouseUp()
Dim i As Long, x As Long, y As Long, Buffer As clsBuffer

    ' find out which button we're clicking
    For i = 1 To 6
        x = GUIWindow(GUI_MENU).x + Buttons(i).x
        y = GUIWindow(GUI_MENU).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            If Buttons(i).state = 2 Then
                ' do stuffs
                Select Case i
                    Case 1
                        ' open window
                        OpenGuiWindow 1
                    Case 2
                        ' open window
                        OpenGuiWindow 2
                    Case 3
                        ' open window
                        OpenGuiWindow 3
                    Case 4
                        ' open window
                        OpenGuiWindow 4
                    Case 5
                        If myTargetType = TARGET_TYPE_PLAYER And myTarget <> MyIndex Then
                            SendTradeRequest
                        Else
                            AddText "Invalid trade target.", BrightRed
                        End If
                    Case 6
                        ' open window
                        OpenGuiWindow 6
                End Select
                ' play sound
                Play_Sound Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub Menu_MouseDown(ByVal Button As Long)
Dim i As Long, x As Long, y As Long
    ' find out which button we're clicking
    For i = 1 To 6
        x = GUIWindow(GUI_MENU).x + Buttons(i).x
        y = GUIWindow(GUI_MENU).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            Buttons(i).state = 2 ' clicked
        End If
    Next
End Sub

' Main Menu
Public Sub MainMenu_MouseUp()
Dim i As Long, x As Long, y As Long, Buffer As clsBuffer

    If faderAlpha > 0 Then Exit Sub

    ' find out which button we're clicking
    For i = 7 To 15
        x = GUIWindow(GUI_MAINMENU).x + Buttons(i).x
        y = GUIWindow(GUI_MAINMENU).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            If Buttons(i).state = 2 Then
                ' do stuffs
                Select Case i
                    Case 7
                        ' login
                        DestroyTCP
                        curMenu = MENU_LOGIN
                        ' clear the textbox
                        sUser = vbNullString
                        sPass = vbNullString
                        curTextbox = 1
                    Case 8
                        ' register
                        DestroyTCP
                        curMenu = MENU_REGISTER
                        ' clear the textbox
                        sUser = vbNullString
                        sPass = vbNullString
                        sPass2 = vbNullString
                        curTextbox = 1
                    Case 9
                        ' credits
                        DestroyTCP
                        curMenu = MENU_CREDITS
                    Case 10
                        ' exit
                        DestroyGame
                        Exit Sub
                    Case 11
                        If curMenu = MENU_LOGIN Then
                            ' login accept
                            MenuState MENU_STATE_LOGIN
                        End If
                    Case 12
                        If curMenu = MENU_REGISTER Then
                            ' register accept
                            MenuState MENU_STATE_NEWACCOUNT
                        End If
                    Case 13
                        If curMenu = MENU_CLASS Then
                            ' they've selected class - move on
                            sChar = vbNullString
                            curMenu = MENU_NEWCHAR
                        End If
                    Case 14
                        If curMenu = MENU_CLASS Then
                            ' next class
                            newCharClass = newCharClass + 1
                            If newCharClass > 3 Then
                                newCharClass = 1
                            End If
                        End If
                    Case 15
                        If curMenu = MENU_NEWCHAR Then
                            ' do eet
                            MenuState MENU_STATE_ADDCHAR
                        End If
                End Select
                ' play sound
                Play_Sound Sound_ButtonClick
            End If
        End If
    Next
    
    ' reset buttons
    resetClickedButtons
End Sub

Public Sub MainMenu_MouseDown(ByVal Button As Long)
Dim i As Long, x As Long, y As Long

    If faderAlpha > 0 Then Exit Sub

    ' find out which button we're clicking
    For i = 7 To 15
        x = GUIWindow(GUI_MAINMENU).x + Buttons(i).x
        y = GUIWindow(GUI_MAINMENU).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            Buttons(i).state = 2 ' clicked
        End If
    Next
End Sub

' Inventory
Public Sub Inventory_MouseUp()
Dim invSlot As Long
    
    If InTrade > 0 Then Exit Sub
    If InBank Or InShop Then Exit Sub

    If DragInvSlotNum > 0 Then
        invSlot = IsInvItem(GlobalX, GlobalY, True)
        If invSlot = 0 Then Exit Sub
        ' change slots
        SendChangeInvSlots DragInvSlotNum, invSlot
    End If

    DragInvSlotNum = 0
End Sub

Public Sub Inventory_MouseDown(ByVal Button As Long)
Dim invNum As Long

    invNum = IsInvItem(GlobalX, GlobalY)

    If Button = 1 Then
        If invNum <> 0 Then
            If InTrade > 0 Then Exit Sub
            If InBank Or InShop Then Exit Sub
            DragInvSlotNum = invNum
        End If

    ElseIf Button = 2 Then
        If Not InBank And Not InShop And Not InTrade > 0 Then
            If invNum <> 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(MyIndex, invNum) > 0 Then
                        CurrencyMenu = 1 ' drop
                        frmMain.lblCurrency.Caption = "How many do you want to drop?"
                        tmpCurrencyItem = invNum
                        frmMain.txtCurrency.Text = vbNullString
                        frmMain.picCurrency.visible = True
                        frmMain.txtCurrency.SetFocus
                    End If
                Else
                    Call SendDropItem(invNum, 0)
                End If
            End If
        End If
    End If
End Sub

Public Sub Inventory_DoubleClick()
    Dim invNum As Long, value As Long, multiplier As Double, i As Long

    DragInvSlotNum = 0
    invNum = IsInvItem(GlobalX, GlobalY)

    If invNum > 0 Then
        ' are we in a shop?
        If InShop > 0 Then
            SellItem invNum
            Exit Sub
        End If
        
        ' in bank?
        If InBank Then
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 2 ' deposit
                frmMain.lblCurrency.Caption = "How many do you want to deposit?"
                tmpCurrencyItem = invNum
                frmMain.txtCurrency.Text = vbNullString
                frmMain.picCurrency.visible = True
                frmMain.txtCurrency.SetFocus
                Exit Sub
            End If
                
            Call DepositItem(invNum, 0)
            Exit Sub
        End If
        
        ' in trade?
        If InTrade > 0 Then
            ' exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = invNum Then
                    ' is currency?
                    If Item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).Type = ITEM_TYPE_CURRENCY Then
                        ' only exit out if we're offering all of it
                        If TradeYourOffer(i).value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then
                            Exit Sub
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_CURRENCY Then
                CurrencyMenu = 4 ' offer in trade
                frmMain.lblCurrency.Caption = "How many do you want to trade?"
                tmpCurrencyItem = invNum
                frmMain.txtCurrency.Text = vbNullString
                frmMain.picCurrency.visible = True
                frmMain.txtCurrency.SetFocus
                Exit Sub
            End If
            
            Call TradeItem(invNum, 0)
            Exit Sub
        End If
        
        ' use item if not doing anything else
        If Item(GetPlayerInvItemNum(MyIndex, invNum)).Type = ITEM_TYPE_NONE Then Exit Sub
        Call SendUseItem(invNum)
        Exit Sub
    End If
End Sub

' Spells
Public Sub Spells_DoubleClick()
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)

    If spellnum <> 0 Then
        Call CastSpell(spellnum)
        Exit Sub
    End If
End Sub

Public Sub Spells_MouseDown(ByVal Button As Long)
Dim spellnum As Long

    spellnum = IsPlayerSpell(GlobalX, GlobalY)
    If Button = 1 Then ' left click
        If spellnum <> 0 Then
            DragSpell = spellnum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' right click
        If spellnum <> 0 Then
            If PlayerSpells(spellnum).Spell > 0 Then
                Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(spellnum).Spell).name) & "?", DIALOGUE_TYPE_FORGET, True, spellnum
            End If
        End If
    End If
End Sub

Public Sub Spells_MouseUp()
Dim spellSlot As Long

    If DragSpell > 0 Then
        spellSlot = IsPlayerSpell(GlobalX, GlobalY, True)
        If spellSlot = 0 Then Exit Sub
        ' drag it
        SendChangeSpellSlots DragSpell, spellSlot
    End If

    DragSpell = 0
End Sub

' character
Public Sub Character_DoubleClick()
Dim eqNum As Long

    eqNum = IsEqItem(GlobalX, GlobalY)

    If eqNum <> 0 Then
        SendUnequip eqNum
    End If
End Sub

Public Sub Character_MouseDown()
Dim i As Long, x As Long, y As Long
    ' find out which button we're clicking
    For i = 16 To 20
        x = GUIWindow(GUI_CHARACTER).x + Buttons(i).x
        y = GUIWindow(GUI_CHARACTER).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            Buttons(i).state = 2 ' clicked
        End If
    Next
End Sub

Public Sub Character_MouseUp()
Dim i As Long, x As Long, y As Long
    ' find out which button we're clicking
    For i = 16 To 20
        x = GUIWindow(GUI_CHARACTER).x + Buttons(i).x
        y = GUIWindow(GUI_CHARACTER).y + Buttons(i).y
        ' check if we're on the button
        If (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            ' send the level up
            If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
            SendTrainStat (i - 15)
            ' play sound
            Play_Sound Sound_ButtonClick
        End If
    Next
End Sub

' hotbar
Public Sub Hotbar_DoubleClick()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarUse slotNum
    End If
End Sub

Public Sub Hotbar_MouseDown(ByVal Button As Long)
Dim slotNum As Long
    
    If Button <> 2 Then Exit Sub ' right click
    
    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum > 0 Then
        SendHotbarChange 0, 0, slotNum
    End If
End Sub

Public Sub Hotbar_MouseUp()
Dim slotNum As Long

    slotNum = IsHotbarSlot(GlobalX, GlobalY)
    If slotNum = 0 Then Exit Sub
    
    ' inventory
    If DragInvSlotNum > 0 Then
        SendHotbarChange 1, DragInvSlotNum, slotNum
        DragInvSlotNum = 0
        Exit Sub
    End If
    
    ' spells
    If DragSpell > 0 Then
        SendHotbarChange 2, DragSpell, slotNum
        DragSpell = 0
        Exit Sub
    End If
End Sub

' Actual input
Public Sub CheckKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetAsyncKeyState(VK_W) >= 0 Then wDown = False
    If GetAsyncKeyState(VK_S) >= 0 Then sDown = False
    If GetAsyncKeyState(VK_A) >= 0 Then aDown = False
    If GetAsyncKeyState(VK_D) >= 0 Then dDown = False
    
    If GetAsyncKeyState(VK_UP) >= 0 Then upDown = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then downDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then leftDown = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then rightDown = False
    
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False
    
    If GetAsyncKeyState(VK_TAB) >= 0 Then tabDown = False
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub CheckInputKeys()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If

    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    
    If GetKeyState(vbKeyTab) < 0 Then
        tabDown = True
    Else
        tabDown = False
    End If

    'Move Up
    If Not chatOn Then
        If GetKeyState(vbKeySpace) < 0 Then
            CheckMapGetItem
        End If
    
        ' move up
        If GetKeyState(vbKeyW) < 0 Then
            wDown = True
            sDown = False
            aDown = False
            dDown = False
            Exit Sub
        Else
            wDown = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyD) < 0 Then
            wDown = False
            sDown = False
            aDown = False
            dDown = True
            Exit Sub
        Else
            dDown = False
        End If
    
        'Move down
        If GetKeyState(vbKeyS) < 0 Then
            wDown = False
            sDown = True
            aDown = False
            dDown = False
            Exit Sub
        Else
            sDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyA) < 0 Then
            wDown = False
            sDown = False
            aDown = True
            dDown = False
            Exit Sub
        Else
            aDown = False
        End If
        
        ' move up
        If GetKeyState(vbKeyUp) < 0 Then
            upDown = True
            leftDown = False
            downDown = False
            rightDown = False
            Exit Sub
        Else
            upDown = False
        End If
    
        'Move Right
        If GetKeyState(vbKeyRight) < 0 Then
            upDown = False
            leftDown = False
            downDown = False
            rightDown = True
            Exit Sub
        Else
            rightDown = False
        End If
    
        'Move down
        If GetKeyState(vbKeyDown) < 0 Then
            upDown = False
            leftDown = False
            downDown = True
            rightDown = False
            Exit Sub
        Else
            downDown = False
        End If
    
        'Move left
        If GetKeyState(vbKeyLeft) < 0 Then
            upDown = False
            leftDown = True
            downDown = False
            rightDown = False
            Exit Sub
        Else
            leftDown = False
        End If
    Else
        wDown = False
        sDown = False
        aDown = False
        dDown = False
        upDown = False
        leftDown = False
        downDown = False
        rightDown = False
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "CheckInputKeys", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub HandleKeyUp(ByVal keyCode As Long)
Dim i As Long

    If InGame Then
        ' admin pannel
        Select Case keyCode
            Case vbKeyInsert
                If Player(MyIndex).Access > 0 Then
                    frmMain.picAdmin.visible = Not frmMain.picAdmin.visible
                End If
        End Select
        
        ' hotbar
        If Not chatOn Then
            For i = 1 To 9
                If keyCode = 48 + i Then
                    SendHotbarUse i
                End If
            Next
            If keyCode = 48 Then ' 0
                SendHotbarUse 10
            ElseIf keyCode = 189 Then ' -
                SendHotbarUse 11
            ElseIf keyCode = 187 Then ' =
                SendHotbarUse 12
            End If
        End If
    End If
    
    ' exit out of fade
    If inMenu Then
        If keyCode = vbKeyEscape Then
            If faderState < 4 Then
                faderState = 4
                faderAlpha = 0
            End If
        End If
    End If
End Sub

Public Sub HandleMenuKeyPresses(ByVal KeyAscii As Integer)
    If Not curMenu = MENU_LOGIN And Not curMenu = MENU_REGISTER And Not curMenu = MENU_NEWCHAR Then Exit Sub
    
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        Select Case curMenu
            Case MENU_LOGIN
                ' next textbox
                If curTextbox = 1 Then
                    curTextbox = 2
                ElseIf curTextbox = 2 Then
                    If KeyAscii = vbKeyTab Then
                        curTextbox = 1
                    Else
                        MenuState MENU_STATE_LOGIN
                    End If
                End If
            Case MENU_REGISTER
                ' next textbox
                If curTextbox = 1 Then
                    curTextbox = 2
                ElseIf curTextbox = 2 Then
                    curTextbox = 3
                ElseIf curTextbox = 3 Then
                    If KeyAscii = vbKeyTab Then
                        curTextbox = 1
                    Else
                        MenuState MENU_STATE_NEWACCOUNT
                    End If
                End If
            Case MENU_NEWCHAR
                If KeyAscii = vbKeyReturn Then
                    MenuState MENU_STATE_ADDCHAR
                End If
        End Select
    End If
    
    Select Case curMenu
        Case MENU_LOGIN
            If curTextbox = 1 Then
                ' entering username
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sUser) > 0 Then sUser = Mid$(sUser, 1, Len(sUser) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sUser = sUser & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 2 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass) > 0 Then sPass = Mid$(sPass, 1, Len(sPass) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sPass = sPass & ChrW$(KeyAscii)
                End If
            End If
        Case MENU_REGISTER
            If curTextbox = 1 Then
                ' entering username
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sUser) > 0 Then sUser = Mid$(sUser, 1, Len(sUser) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sUser = sUser & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 2 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass) > 0 Then sPass = Mid$(sPass, 1, Len(sPass) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sPass = sPass & ChrW$(KeyAscii)
                End If
            ElseIf curTextbox = 3 Then
                ' entering password
                If (KeyAscii = vbKeyBack) Then
                    If LenB(sPass2) > 0 Then sPass2 = Mid$(sPass2, 1, Len(sPass2) - 1)
                End If
            
                ' And if neither, then add the character to the user's text buffer
                If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                    sPass2 = sPass2 & ChrW$(KeyAscii)
                End If
            End If
        Case MENU_NEWCHAR
            ' entering username
            If (KeyAscii = vbKeyBack) Then
                If LenB(sChar) > 0 Then sChar = Mid$(sChar, 1, Len(sChar) - 1)
            End If
        
            ' And if neither, then add the character to the user's text buffer
            If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) And (KeyAscii <> vbKeyTab) Then
                sChar = sChar & ChrW$(KeyAscii)
            End If
    End Select
End Sub

Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim chatText As String
Dim name As String
Dim i As Long
Dim n As Long
Dim Command() As String
Dim Buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    chatText = MyText

    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
    
        ' turn on/off the chat
        chatOn = Not chatOn

        ' Broadcast message
        If left$(chatText, 1) = "'" Then
            chatText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call BroadcastMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Emote message
        If left$(chatText, 1) = "-" Then
            MyText = Mid$(chatText, 2, Len(chatText) - 1)

            If Len(chatText) > 0 Then
                Call EmoteMsg(chatText)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Player message
        If left$(chatText, 1) = "!" Then
            Exit Sub
            chatText = Mid$(chatText, 2, Len(chatText) - 1)
            name = vbNullString

            ' Get the desired player from the user text
            For i = 1 To Len(chatText)

                If Mid$(chatText, i, 1) <> Space(1) Then
                    name = name & Mid$(chatText, i, 1)
                Else
                    Exit For
                End If

            Next

            chatText = Mid$(chatText, i, Len(chatText) - 1)

            ' Make sure they are actually sending something
            If Len(chatText) - i > 0 Then
                MyText = Mid$(chatText, i + 1, Len(chatText) - i)
                ' Send the message to the player
                Call PlayerMsg(chatText, name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If

            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        If left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))

            Select Case Command(0)
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Global Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /who, /fps, /fpslock, /gui, /maps", HelpColor)
                Case "/maps"
                    ClearMapCache
                Case "/gui"
                    hideGUI = Not hideGUI
                Case "/info"

                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo continue
                    End If

                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CPlayerInfoRequest
                    Buffer.WriteString Command(1)
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                    ' Whos Online
                Case "/who"
                    SendWhosOnline
                    ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    ' toggle fps lock
                Case "/fpslock"
                    FPS_Lock = Not FPS_Lock
                    ' Request stats
                Case "/stats"
                    Set Buffer = New clsBuffer
                    Buffer.WriteLong CGetStats
                    SendData Buffer.ToArray()
                    Set Buffer = Nothing
                    ' // Monitor Admin Commands //
                    ' Admin Help
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo continue
                    frmMain.picAdmin.visible = Not frmMain.picAdmin.visible
                    ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo continue
                    End If

                    SendKick Command(1)
                    ' // Mapper Admin Commands //
                    ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    BLoc = Not BLoc
                    ' Map Editor
                Case "/editmap"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue
                    
                    SendRequestEditMap
                    ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo continue
                    End If
                    
                    GettingMap = True
                    WarpMeTo Command(1)
                    ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo continue
                    End If
                    
                    WarpToMe Command(1)
                    ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo continue
                    End If

                    n = CLng(Command(1))

                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        GettingMap = True
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If
                    ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo continue
                    End If

                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo continue
                    End If

                    SendSetSprite CLng(Command(1))
                    ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    SendMapReport
                    ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    SendMapRespawn
                    ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo continue
                    End If

                    SendMOTDChange Right$(chatText, Len(chatText) - 5)
                    ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    SendBanList
                    ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then GoTo continue

                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo continue
                    End If

                    SendBan Command(1)
                    ' // Developer Admin Commands //
                    ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditItem
                ' editing conv request
                Case "/editconv"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditConv
                ' Editing animation request
                Case "/editanimation"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditAnimation
                    ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditNpc
                Case "/editresource"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditResource
                    ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditShop
                    ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then GoTo continue

                    SendRequestEditSpell
                    ' // Creator Admin Commands //
                    ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue

                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo continue
                    End If

                    SendSetAccess Command(1), CLng(Command(2))
                    ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue

                    SendBanDestroy
                    ' Packet debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then GoTo continue

                    DEBUG_MODE = (Not DEBUG_MODE)
                Case Else
                    AddText "Not a valid command!", HelpColor
            End Select

            'continue label where we go instead of exiting the sub
continue:
            MyText = vbNullString
            UpdateShowChatText
            Exit Sub
        End If

        ' Say message
        If Len(chatText) > 0 Then
            Call SayMsg(MyText)
        End If

        MyText = vbNullString
        UpdateShowChatText
        Exit Sub
    End If
    
    If Not chatOn Then Exit Sub

    ' Handle when the user presses the backspace key
    If (KeyAscii = vbKeyBack) Then
        If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        UpdateShowChatText
    End If

    ' And if neither, then add the character to the user's text buffer
    If (KeyAscii <> vbKeyReturn) Then
        If (KeyAscii <> vbKeyBack) Then
            MyText = MyText & ChrW$(KeyAscii)
            UpdateShowChatText
        End If
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "HandleKeyPresses", "modInput", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub
