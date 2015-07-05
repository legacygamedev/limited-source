Attribute VB_Name = "modHandleData"
Option Explicit

Sub HandleData(ByVal Data As String)
    Dim parse() As String
    Dim Name As String
    Dim Msg As String
    Dim Dir As Long
    Dim Level As Long
    Dim i As Long, n As Long, X As Long, y As Long, p As Long
    Dim shopNum As Long
    Dim z As Long
    Dim strfilename As String
    Dim CustomX As Long
    Dim CustomY As Long
    Dim CustomIndex As Long
    Dim customcolour As Long
    Dim customsize As Long
    Dim customtext As String
    Dim casestring As String
    Dim packet As String
    Dim m As Long
    Dim j As Long

    ' Handle Data
    parse = Split(Data, SEP_CHAR)

    ' Add packet info to debugger
    If frmDebug.Visible = True Then
        Call TextAdd(frmDebug.txtDebug, Time & " - ( " & parse(0) & " )", True)
    End If

' :::::::::::::::::::::::
' :: Get players stats ::
' :::::::::::::::::::::::

    casestring = LCase$(parse(0))

    If casestring = "leaveparty211" Then
        For i = 1 To MAX_PARTY_MEMBERS
            Player(MyIndex).Party.Member(i) = 0
        Next i
        Exit Sub
    End If

    If casestring = "playerhpreturn" Then
        Player(Val(parse(1))).HP = Val(parse(2))
        Player(Val(parse(1))).MaxHp = Val(parse(3))
        ' Call MsgBox("player(" & val(parse(1)) & ").hp = " & val(parse(2)))
        ' Call BltPlayerBars(val(parse(1)))
        Exit Sub
    End If

    If casestring = "maxinfo" Then
        GAME_NAME = Trim$(parse(1))
        MAX_PLAYERS = Val(parse(2))
        MAX_ITEMS = Val(parse(3))
        MAX_NPCS = Val(parse(4))
        MAX_SHOPS = Val(parse(5))
        MAX_SPELLS = Val(parse(6))
        MAX_MAPS = Val(parse(7))
        MAX_MAP_ITEMS = Val(parse(8))
        MAX_MAPX = Val(parse(9))
        MAX_MAPY = Val(parse(10))
        MAX_EMOTICONS = Val(parse(11))
        MAX_ELEMENTS = Val(parse(12))
        paperdoll = Val(parse(13))
        SpriteSize = Val(parse(14))
        MAX_SCRIPTSPELLS = Val(parse(15))
        CustomPlayers = Val(parse(16))
        lvl = Val(parse(17))
        MAX_PARTY_MEMBERS = Val(parse(18))
        STAT1 = parse(19)
        STAT2 = parse(20)
        STAT3 = parse(21)
        STAT4 = parse(22)

        If 0 + CustomPlayers > 0 Then
            frmNewChar.Picture4.Visible = False
            frmNewChar.HScroll1.Visible = True
            frmNewChar.HScroll2.Visible = True
            frmNewChar.HScroll3.Visible = True
            frmNewChar.Label14.Visible = True
            frmNewChar.Label11.Visible = True
            frmNewChar.Label12.Visible = True
            frmNewChar.Picture1.Visible = True

            If FileExists("GFX\Heads.bmp") Then
                frmNewChar.iconn(0).Picture = LoadPicture(App.Path & "\GFX\Heads.bmp")
            End If
            If FileExists("GFX\Bodys.bmp") Then
                frmNewChar.iconn(1).Picture = LoadPicture(App.Path & "\GFX\Bodys.bmp")
            End If
            If FileExists("GFX\Legs.bmp") Then
                frmNewChar.iconn(2).Picture = LoadPicture(App.Path & "\GFX\Legs.bmp")
            End If


            If SpriteSize = 1 Then
                frmNewChar.iconn(0).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(0).Top = -Val(PIC_Y - 15)

                frmNewChar.iconn(1).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(1).Top = -Val(PIC_Y - 7)

                frmNewChar.iconn(2).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(2).Top = -Val(PIC_Y + 3)
            Else
                frmNewChar.iconn(0).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(0).Top = -Val(PIC_Y)

                frmNewChar.iconn(1).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(1).Top = -Val(PIC_Y)

                frmNewChar.iconn(2).Left = -Val(5 * PIC_X)
                frmNewChar.iconn(2).Top = -Val(PIC_Y)
            End If
        End If

        ReDim Map(0 To MAX_MAPS) As MapRec
        ReDim Player(1 To MAX_PLAYERS) As PlayerRec
        ReDim Item(1 To MAX_ITEMS) As ItemRec
        ReDim Npc(1 To MAX_NPCS) As NpcRec
        ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim Shop(1 To MAX_SHOPS) As ShopRec
        ReDim Spell(1 To MAX_SPELLS) As SpellRec
        ReDim Element(0 To MAX_ELEMENTS) As ElementRec
        ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
        ReDim ScriptBubble(1 To MAX_BUBBLES) As ScriptBubble
        ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim ScriptSpell(1 To MAX_SCRIPTSPELLS) As ScriptSpellAnimRec

        For i = 1 To MAX_MAPS
            ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
            ReDim Map(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        Next i
        
        ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
        ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
        ReDim MapReport(1 To MAX_MAPS) As MapRec
        
        MAX_SPELL_ANIM = MAX_MAPX * MAX_MAPY

        MAX_BLT_LINE = 6
        ReDim BattlePMsg(1 To MAX_BLT_LINE) As BattleMsgRec
        ReDim BattleMMsg(1 To MAX_BLT_LINE) As BattleMsgRec

        For i = 1 To MAX_PLAYERS
            ReDim Player(i).SpellAnim(1 To MAX_SPELL_ANIM) As SpellAnimRec
        Next i

        For i = 0 To MAX_EMOTICONS
            Emoticons(i).Pic = 0
            Emoticons(i).Command = vbNullString
        Next i

        Call ClearTempTile

        ' Clear out players
        For i = 1 To MAX_PLAYERS
            Call ClearPlayer(i)
        Next i

        For i = 1 To MAX_MAPS
            Call LoadMap(i)
        Next i

        frmMirage.Caption = Trim$(GAME_NAME)
        App.Title = GAME_NAME

        AllDataReceived = True

        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If casestring = "npchp" Then
        n = Val(parse(1))

        MapNpc(n).HP = Val(parse(2))
        MapNpc(n).MaxHp = Val(parse(3))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "alertmsg" Then
        frmMirage.Visible = False
        frmSendGetData.Visible = False
        frmMainMenu.Visible = True

        Msg = parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Plain message packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "plainmsg" Then
        frmSendGetData.Visible = False
        n = Val(parse(2))

        If n = 0 Then
            frmMainMenu.Show
        End If
        If n = 1 Then
            frmNewAccount.Show
        End If
        If n = 2 Then
            frmDeleteAccount.Show
        End If
        If n = 3 Then
            frmLogin.Show
        End If
        If n = 4 Then
            frmNewChar.Show
        End If
        If n = 5 Then
            frmChars.Show
        End If

        Msg = parse(1)
        Call MsgBox(Msg, vbOKOnly, GAME_NAME)
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If casestring = "allchars" Then

        n = 1

        frmChars.Visible = True
        frmSendGetData.Visible = False

        frmChars.lstChars.Clear

        For i = 1 To MAX_CHARS
            Name = parse(n)
            Msg = parse(n + 1)
            Level = Val(parse(n + 2))

            If Trim$(Name) = vbNullString Then
                frmChars.lstChars.addItem "Free Character Slot"
            Else
                frmChars.lstChars.addItem Name & " a level " & Level & " " & Msg
            End If

            n = n + 3
        Next i

        frmChars.lstChars.ListIndex = 0

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If casestring = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(parse(1))

        frmSendGetData.Visible = True
        frmChars.Visible = False

        ReDim Player(MyIndex).Party.Member(1 To MAX_PARTY_MEMBERS)

        Call SetStatus("Receiving game data...")
        Exit Sub
    End If


    ' :::::::::::::::::::::::::::::::::
    ' ::     News Recieved packet    ::
    ' :::::::::::::::::::::::::::::::::
    If casestring = "news" Then
        Call WriteINI("DATA", "News", parse(1), (App.Path & "\News.ini"))
        Call WriteINI("DATA", "Desc", parse(5), (App.Path & "\News.ini"))
        Call WriteINI("COLOR", "Red", CInt(parse(2)), (App.Path & "\News.ini"))
        Call WriteINI("COLOR", "Green", CInt(parse(3)), (App.Path & "\News.ini"))
        Call WriteINI("COLOR", "Blue", CInt(parse(4)), (App.Path & "\News.ini"))

        ' We just gots teh news, so change the news label
        Call ParseNews
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "newcharclasses" Then
        n = 1

        ' ClassesOn
        ClassesOn = Int(parse(2))

        ' Max classes
        Max_Classes = Val(parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec

        n = n + 2

        For i = 0 To Max_Classes
            Class(i).Name = parse(n)

            Class(i).HP = Val(parse(n + 1))
            Class(i).MP = Val(parse(n + 2))
            Class(i).SP = Val(parse(n + 3))

            Class(i).STR = Val(parse(n + 4))
            Class(i).DEF = Val(parse(n + 5))
            Class(i).speed = Val(parse(n + 6))
            Class(i).MAGI = Val(parse(n + 7))
            ' Class(i).INTEL = val(parse(n + 8))
            Class(i).MaleSprite = Val(parse(n + 8))
            Class(i).FemaleSprite = Val(parse(n + 9))
            Class(i).Locked = Val(parse(n + 10))
            Class(i).desc = parse(n + 11)

            n = n + 12
        Next i


        ' Used for if the player is creating a new character
        frmNewChar.Visible = True
        frmSendGetData.Visible = False

        frmNewChar.cmbClass.Clear
        For i = 0 To Max_Classes
            If Class(i).Locked = 0 Then
                frmNewChar.cmbClass.addItem Trim$(Class(i).Name)
            End If
        Next i
        frmNewChar.cmbClass.ListIndex = 0
        frmNewChar.lblClassDesc = Class(0).desc
        If ClassesOn = 1 Then
            frmNewChar.cmbClass.Visible = True
            frmNewChar.lblClassDesc.Visible = True
        ElseIf ClassesOn = 0 Then
            frmNewChar.cmbClass.Visible = False
            frmNewChar.lblClassDesc.Visible = False
        End If


        frmNewChar.lblHP.Caption = STR(Class(0).HP)
        frmNewChar.lblMP.Caption = STR(Class(0).MP)
        frmNewChar.lblSP.Caption = STR(Class(0).SP)

        frmNewChar.lblSTR.Caption = STR(Class(0).STR)
        frmNewChar.lblDEF.Caption = STR(Class(0).DEF)
        frmNewChar.lblSPEED.Caption = STR(Class(0).speed)
        frmNewChar.lblMAGI.Caption = STR(Class(0).MAGI)

        frmNewChar.lblClassDesc.Caption = Class(0).desc
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "classesdata" Then
        n = 1

        ' Max classes
        Max_Classes = Val(parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec

        n = n + 1

        For i = 0 To Max_Classes
            Class(i).Name = parse(n)

            Class(i).HP = Val(parse(n + 1))
            Class(i).MP = Val(parse(n + 2))
            Class(i).SP = Val(parse(n + 3))

            Class(i).STR = Val(parse(n + 4))
            Class(i).DEF = Val(parse(n + 5))
            Class(i).speed = Val(parse(n + 6))
            Class(i).MAGI = Val(parse(n + 7))

            Class(i).Locked = Val(parse(n + 8))
            Class(i).desc = parse(n + 9)

            n = n + 10
        Next i
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' ::  Game Clock (Time)  ::
    ' :::::::::::::::::::::::::
    If casestring = "gameclock" Then
        Seconds = Val(parse(1))
        Minutes = Val(parse(2))
        Hours = Val(parse(3))
        Gamespeed = Val(parse(4))
        frmMirage.lblGameTime.Caption = "It is now:"
        frmMirage.lblGameTime.Visible = True
    End If

    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If casestring = "ingame" Then
        Call GameInit
        Call GameLoop
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If casestring = "playerinv" Then
        n = 2
        z = Val(parse(1))

        For i = 1 To MAX_INV
            Call SetPlayerInvItemNum(z, i, Val(parse(n)))
            Call SetPlayerInvItemValue(z, i, Val(parse(n + 1)))
            Call SetPlayerInvItemDur(z, i, Val(parse(n + 2)))

            n = n + 3
        Next i

        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If casestring = "playerinvupdate" Then
        n = Val(parse(1))
        z = Val(parse(2))

        Call SetPlayerInvItemNum(z, n, Val(parse(3)))
        Call SetPlayerInvItemValue(z, n, Val(parse(4)))
        Call SetPlayerInvItemDur(z, n, Val(parse(5)))
        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If
    ' ::::::::::::::::::::::::
    ' :: Player bank packet ::
    ' ::::::::::::::::::::::::
    If casestring = "playerbank" Then
        n = 1
        For i = 1 To MAX_BANK
            Call SetPlayerBankItemNum(MyIndex, i, Val(parse(n)))
            Call SetPlayerBankItemValue(MyIndex, i, Val(parse(n + 1)))
            Call SetPlayerBankItemDur(MyIndex, i, Val(parse(n + 2)))

            n = n + 3
        Next i

        If frmBank.Visible = True Then
            Call UpdateBank
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Player bank update packet ::
    ' :::::::::::::::::::::::::::::::
    If casestring = "playerbankupdate" Then
        n = Val(parse(1))

        Call SetPlayerBankItemNum(MyIndex, n, Val(parse(2)))
        Call SetPlayerBankItemValue(MyIndex, n, Val(parse(3)))
        Call SetPlayerBankItemDur(MyIndex, n, Val(parse(4)))
        If frmBank.Visible = True Then
            Call UpdateBank
        End If
        Exit Sub
    End If

' :::::::::::::::::::::::::::::::
' :: Player bank open packet ::
' :::::::::::::::::::::::::::::::

    If casestring = "openbank" Then
        ' frmBank.lblBank.Caption = Trim$(Map(GetPlayerMap(MyIndex)).Name)
        frmBank.lstInventory.Clear
        frmBank.lstBank.Clear
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerInvItemNum(MyIndex, i)).Stackable = 1 Then
                    frmBank.lstInventory.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                        frmBank.lstInventory.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                    Else
                        frmBank.lstInventory.addItem i & "> " & Trim$(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                    End If
                End If
            Else
                frmBank.lstInventory.addItem i & "> Empty"
            End If

        Next i

        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(MyIndex, i) > 0 Then
                If Item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Or Item(GetPlayerBankItemNum(MyIndex, i)).Stackable = 1 Then
                    frmBank.lstBank.addItem i & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerLegsSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerNecklaceSlot(MyIndex) = i Then
                        frmBank.lstBank.addItem i & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (worn)"
                    Else
                        frmBank.lstBank.addItem i & "> " & Trim$(Item(GetPlayerBankItemNum(MyIndex, i)).Name)
                    End If
                End If
            Else
                frmBank.lstBank.addItem i & "> Empty"
            End If

        Next i
        frmBank.lstBank.ListIndex = 0
        frmBank.lstInventory.ListIndex = 0

        frmBank.Show vbModal
        Exit Sub
    End If

    If LCase$(parse(0)) = "bankmsg" Then
        frmBank.lblMsg.Caption = Trim$(parse(1))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If casestring = "playerworneq" Then

        z = Val(parse(1))
        If z <= 0 Then
            Exit Sub
        End If
        Call SetPlayerArmorSlot(z, Val(parse(2)))
        Call SetPlayerWeaponSlot(z, Val(parse(3)))
        Call SetPlayerHelmetSlot(z, Val(parse(4)))
        Call SetPlayerShieldSlot(z, Val(parse(5)))
        Call SetPlayerLegsSlot(z, Val(parse(6)))
        Call SetPlayerRingSlot(z, Val(parse(7)))
        Call SetPlayerNecklaceSlot(z, Val(parse(8)))

        If z = MyIndex Then
            Call UpdateVisInv
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Player points packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "playerpoints" Then
        Player(MyIndex).POINTS = Val(parse(1))

        If GetPlayerPOINTS(MyIndex) > 0 Then
            frmMirage.AddSTR.Visible = True
            frmMirage.AddDEF.Visible = True
            frmMirage.AddSPD.Visible = True
            frmMirage.AddMAGI.Visible = True
        Else
            frmMirage.AddSTR.Visible = False
            frmMirage.AddDEF.Visible = False
            frmMirage.AddSPD.Visible = False
            frmMirage.AddMAGI.Visible = False
        End If

        frmMirage.lblPoints.Caption = Val(parse(1))
        Exit Sub
    End If

    If casestring = "cussprite" Then
        Player(Val(parse(1))).head = Val(parse(2))
        Player(Val(parse(1))).body = Val(parse(3))
        Player(Val(parse(1))).leg = Val(parse(4))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase$(casestring) = "playerhp" Then
        Player(MyIndex).MaxHp = Val(parse(1))
        Call SetPlayerHP(MyIndex, Val(parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            ' frmMirage.shpHP.FillColor = RGB(208, 11, 0)
            frmMirage.shpHP.Width = ((GetPlayerHP(MyIndex)) / (GetPlayerMaxHP(MyIndex))) * 150
            frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
        End If
        Exit Sub
    End If

    If casestring = "playerexp" Then
        Call SetPlayerExp(MyIndex, Val(parse(1)))
        frmMirage.lblEXP.Caption = Val(parse(1)) & " / " & Val(parse(2))
        frmMirage.shpTNL.Width = (((Val(parse(1))) / (Val(parse(2)))) * 150)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If casestring = "playermp" Then
        Player(MyIndex).MaxMP = Val(parse(1))
        Call SetPlayerMP(MyIndex, Val(parse(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            ' frmMirage.shpMP.FillColor = RGB(208, 11, 0)
            frmMirage.shpMP.Width = ((GetPlayerMP(MyIndex)) / (GetPlayerMaxMP(MyIndex))) * 150
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
        End If
        Exit Sub
    End If

    ' speech bubble parse
    If (casestring = "mapmsg2") Then
        Bubble(Val(parse(2))).Text = parse(1)
        Bubble(Val(parse(2))).Created = GetTickCount()
        Exit Sub
    End If

    ' scriptbubble parse
    If (casestring = "scriptbubble") Then
        ScriptBubble(Val(parse(1))).Text = Trim$(parse(2))
        ScriptBubble(Val(parse(1))).Map = Val(parse(3))
        ScriptBubble(Val(parse(1))).X = Val(parse(4))
        ScriptBubble(Val(parse(1))).y = Val(parse(5))
        ScriptBubble(Val(parse(1))).Colour = Val(parse(6))
        ScriptBubble(Val(parse(1))).Created = GetTickCount()
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If casestring = "playersp" Then
        Player(MyIndex).MaxSP = Val(parse(1))
        Call SetPlayerSP(MyIndex, Val(parse(2)))
        If GetPlayerMaxSP(MyIndex) > 0 Then
            frmMirage.lblSP.Caption = Int(GetPlayerSP(MyIndex) / GetPlayerMaxSP(MyIndex) * 100) & "%"
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "playerstatspacket") Then
        Dim SubDef As Long, SubMagi As Long, SubSpeed As Long, SubStr As Long
        SubStr = 0
        SubDef = 0
        SubMagi = 0
        SubSpeed = 0

        If GetPlayerWeaponSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSTR
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerArmorSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSTR
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerShieldSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSTR
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerHelmetSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSTR
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerLegsSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSTR
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerLegsSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerRingSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSTR
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSpeed
        End If
        If GetPlayerNecklaceSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSTR
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddDEF
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddMAGI
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerNecklaceSlot(MyIndex))).AddSpeed
        End If

        If SubStr > 0 Then
            frmMirage.lblSTR.Caption = Val(parse(1)) - SubStr & " (+" & SubStr & ")"
        Else
            frmMirage.lblSTR.Caption = Val(parse(1))
        End If
        If SubDef > 0 Then
            frmMirage.lblDEF.Caption = Val(parse(2)) - SubDef & " (+" & SubDef & ")"
        Else
            frmMirage.lblDEF.Caption = Val(parse(2))
        End If
        If SubMagi > 0 Then
            frmMirage.lblMAGI.Caption = Val(parse(4)) - SubMagi & " (+" & SubMagi & ")"
        Else
            frmMirage.lblMAGI.Caption = Val(parse(4))
        End If
        If SubSpeed > 0 Then
            frmMirage.lblSPEED.Caption = Val(parse(3)) - SubSpeed & " (+" & SubSpeed & ")"
        Else
            frmMirage.lblSPEED.Caption = Val(parse(3))
        End If
        frmMirage.lblEXP.Caption = Val(parse(6)) & " / " & Val(parse(5))

        frmMirage.shpTNL.Width = (((Val(parse(6))) / (Val(parse(5)))) * 150)
        frmMirage.lblLevel.Caption = Val(parse(7))
        Player(MyIndex).Level = Val(parse(7))

        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If casestring = "playerdata" Then
        Dim a As Long
        i = Val(parse(1))
        Call SetPlayerName(i, parse(2))
        Call SetPlayerSprite(i, Val(parse(3)))
        Call SetPlayerMap(i, Val(parse(4)))
        Call SetPlayerX(i, Val(parse(5)))
        Call SetPlayerY(i, Val(parse(6)))
        Call SetPlayerDir(i, Val(parse(7)))
        Call SetPlayerAccess(i, Val(parse(8)))
        Call SetPlayerPK(i, Val(parse(9)))
        Call SetPlayerGuild(i, parse(10))
        Call SetPlayerGuildAccess(i, Val(parse(11)))
        Call SetPlayerClass(i, Val(parse(12)))
        Call SetPlayerHead(i, Val(parse(13)))
        Call SetPlayerBody(i, Val(parse(14)))
        Call SetPlayerLeg(i, Val(parse(15)))
        Call SetPlayerPaperdoll(i, Val(parse(16)))
        Call SetPlayerLevel(i, Val(parse(17)))

        ' Make sure they aren't walking
        Player(i).Moving = 0
        Player(i).xOffset = 0
        Player(i).yOffset = 0

        ' Check if the player is the client player, and if so reset directions
        If i = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        
        
        Exit Sub
        
    End If
    
    ' if a player leaves the map
    If casestring = "leave" Then
        Call SetPlayerMap(CLng(parse(1)), 0)
        Exit Sub
    End If
        
    ' if a player left the game
    If casestring = "left" Then
        Call ClearPlayer(parse(1))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Player Level Packet  ::
    ' ::::::::::::::::::::::::::
    If casestring = "playerlevel" Then
        n = Val(parse(1))
        Player(n).Level = Val(parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Update Sprite Packet ::
    ' ::::::::::::::::::::::::::
    If casestring = "updatesprite" Then
        i = Val(parse(1))
        Call SetPlayerSprite(i, Val(parse(1)))
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "playermove") Then
        i = Val(parse(1))
        X = Val(parse(2))
        y = Val(parse(3))
        Dir = Val(parse(4))
        n = Val(parse(5))

        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Sub
        End If

        Call SetPlayerX(i, X)
        Call SetPlayerY(i, y)
        Call SetPlayerDir(i, Dir)

        Player(i).xOffset = 0
        Player(i).yOffset = 0
        Player(i).Moving = n
        
        ' Replaced with the one from TE.
        Select Case GetPlayerDir(i)
            Case DIR_UP
                Player(i).yOffset = PIC_Y
            Case DIR_DOWN
                Player(i).yOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(i).xOffset = PIC_X
            Case DIR_RIGHT
                Player(i).xOffset = PIC_X * -1
        End Select

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "npcmove") Then
        i = Val(parse(1))
        X = Val(parse(2))
        y = Val(parse(3))
        Dir = Val(parse(4))
        n = Val(parse(5))

        MapNpc(i).X = X
        MapNpc(i).y = y
        MapNpc(i).Dir = Dir
        MapNpc(i).xOffset = 0
        MapNpc(i).yOffset = 0
        MapNpc(i).Moving = 1

        If n <> 1 Then
            Select Case MapNpc(i).Dir
                Case DIR_UP
                    MapNpc(i).yOffset = PIC_Y * Val(n - 1)
                Case DIR_DOWN
                    MapNpc(i).yOffset = PIC_Y * -n
                Case DIR_LEFT
                    MapNpc(i).xOffset = PIC_X * Val(n - 1)
                Case DIR_RIGHT
                    MapNpc(i).xOffset = PIC_X * -n
            End Select
        Else
            Select Case MapNpc(i).Dir
                Case DIR_UP
                    MapNpc(i).yOffset = PIC_Y
                Case DIR_DOWN
                    MapNpc(i).yOffset = PIC_Y * -1
                Case DIR_LEFT
                    MapNpc(i).xOffset = PIC_X
                Case DIR_RIGHT
                    MapNpc(i).xOffset = PIC_X * -1
            End Select
        End If

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (casestring = "playerdir") Then
        i = Val(parse(1))
        Dir = Val(parse(2))

        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Exit Sub
        End If

        Call SetPlayerDir(i, Dir)

        Player(i).xOffset = 0
        Player(i).yOffset = 0
        Player(i).MovingH = 0
        Player(i).MovingV = 0
        Player(i).Moving = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "npcdir") Then
        i = Val(parse(1))
        Dir = Val(parse(2))
        MapNpc(i).Dir = Dir

        MapNpc(i).xOffset = 0
        MapNpc(i).yOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (casestring = "playerxy") Then
        i = Val(parse(1))
        X = Val(parse(2))
        y = Val(parse(3))

        Call SetPlayerX(i, X)
        Call SetPlayerY(i, y)

        ' Make sure they aren't walking
        Player(i).Moving = 0
        Player(i).xOffset = 0
        Player(i).yOffset = 0

        Exit Sub
    End If

    If LCase$(parse(0)) = "removemembers" Then
        For n = 1 To MAX_PARTY_MEMBERS
            Player(MyIndex).Party.Member(n) = 0
        Next n
        Exit Sub
    End If

    If LCase$(parse(0)) = "updatemembers" Then
        Player(MyIndex).Party.Member(Val(parse(1))) = Val(parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "attack") Then
        i = Val(parse(1))

        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount

        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (casestring = "npcattack") Then
        i = Val(parse(1))

        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (casestring = "checkformap") Then
        GettingMap = True
    
        ' Erase all players except self
        For i = 1 To MAX_PLAYERS
            If i <> MyIndex Then
                Call SetPlayerMap(i, 0)
            End If
        Next i

        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        X = Val(parse(1))

        ' Get revision
        y = Val(parse(2))
        
        ' Close map editor if player leaves current map
        If InEditor Then
            ScreenMode = 0
            NightMode = 0
            GridMode = 0
            InEditor = False
            Unload frmMapEditor
            frmMapEditor.MousePointer = 1
            frmMirage.MousePointer = 1
        End If
        

        If FileExists("maps\map" & X & ".dat") Then
            ' Check to see if the revisions match
            If GetMapRevision(X) = y Then
            ' We do so we dont need the map

                ' Load the map
                Call LoadMap(X)

                Call SendData("needmap" & SEP_CHAR & "no" & END_CHAR)
                Exit Sub
            End If
        End If

        ' Either the revisions didn't match or we dont have the map, so we need it
        Call SendData("needmap" & SEP_CHAR & "yes" & END_CHAR)
        Exit Sub
    End If

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::

    If casestring = "mapdata" Then
        n = 1

        Map(Val(parse(1))).Name = parse(n + 1)
        Map(Val(parse(1))).Revision = Val(parse(n + 2))
        Map(Val(parse(1))).Moral = Val(parse(n + 3))
        Map(Val(parse(1))).Up = Val(parse(n + 4))
        Map(Val(parse(1))).Down = Val(parse(n + 5))
        Map(Val(parse(1))).Left = Val(parse(n + 6))
        Map(Val(parse(1))).Right = Val(parse(n + 7))
        Map(Val(parse(1))).music = parse(n + 8)
        Map(Val(parse(1))).BootMap = Val(parse(n + 9))
        Map(Val(parse(1))).BootX = Val(parse(n + 10))
        Map(Val(parse(1))).BootY = Val(parse(n + 11))
        Map(Val(parse(1))).Indoors = Val(parse(n + 12))
        Map(Val(parse(1))).Weather = Val(parse(n + 13))

        n = n + 14

        For y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                Map(Val(parse(1))).Tile(X, y).Ground = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).Mask = Val(parse(n + 1))
                Map(Val(parse(1))).Tile(X, y).Anim = Val(parse(n + 2))
                Map(Val(parse(1))).Tile(X, y).Mask2 = Val(parse(n + 3))
                Map(Val(parse(1))).Tile(X, y).M2Anim = Val(parse(n + 4))
                Map(Val(parse(1))).Tile(X, y).Fringe = Val(parse(n + 5))
                Map(Val(parse(1))).Tile(X, y).FAnim = Val(parse(n + 6))
                Map(Val(parse(1))).Tile(X, y).Fringe2 = Val(parse(n + 7))
                Map(Val(parse(1))).Tile(X, y).F2Anim = Val(parse(n + 8))
                Map(Val(parse(1))).Tile(X, y).Type = Val(parse(n + 9))
                Map(Val(parse(1))).Tile(X, y).Data1 = Val(parse(n + 10))
                Map(Val(parse(1))).Tile(X, y).Data2 = Val(parse(n + 11))
                Map(Val(parse(1))).Tile(X, y).Data3 = Val(parse(n + 12))
                Map(Val(parse(1))).Tile(X, y).String1 = parse(n + 13)
                Map(Val(parse(1))).Tile(X, y).String2 = parse(n + 14)
                Map(Val(parse(1))).Tile(X, y).String3 = parse(n + 15)
                Map(Val(parse(1))).Tile(X, y).light = Val(parse(n + 16))
                Map(Val(parse(1))).Tile(X, y).GroundSet = Val(parse(n + 17))
                Map(Val(parse(1))).Tile(X, y).MaskSet = Val(parse(n + 18))
                Map(Val(parse(1))).Tile(X, y).AnimSet = Val(parse(n + 19))
                Map(Val(parse(1))).Tile(X, y).Mask2Set = Val(parse(n + 20))
                Map(Val(parse(1))).Tile(X, y).M2AnimSet = Val(parse(n + 21))
                Map(Val(parse(1))).Tile(X, y).FringeSet = Val(parse(n + 22))
                Map(Val(parse(1))).Tile(X, y).FAnimSet = Val(parse(n + 23))
                Map(Val(parse(1))).Tile(X, y).Fringe2Set = Val(parse(n + 24))
                Map(Val(parse(1))).Tile(X, y).F2AnimSet = Val(parse(n + 25))
                n = n + 26
            Next X
        Next y

        For X = 1 To 15
            Map(Val(parse(1))).Npc(X) = Val(parse(n))
            Map(Val(parse(1))).SpawnX(X) = Val(parse(n + 1))
            Map(Val(parse(1))).SpawnY(X) = Val(parse(n + 2))
            n = n + 3
        Next X

        ' Save the map
        Call SaveLocalMap(Val(parse(1)))

        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            frmMapEditor.Visible = False
            frmMirage.Show
' frmMirage.picMapEditor.Visible = False

            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If

            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If

        Exit Sub
    End If

    If casestring = "tilecheck" Then
        n = 5
        X = Val(parse(2))
        y = Val(parse(3))

        Select Case Val(parse(4))
            Case 0
                Map(Val(parse(1))).Tile(X, y).Ground = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).GroundSet = Val(parse(n + 1))
            Case 1
                Map(Val(parse(1))).Tile(X, y).Mask = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).MaskSet = Val(parse(n + 1))
            Case 2
                Map(Val(parse(1))).Tile(X, y).Anim = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).AnimSet = Val(parse(n + 1))
            Case 3
                Map(Val(parse(1))).Tile(X, y).Mask2 = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).Mask2Set = Val(parse(n + 1))
            Case 4
                Map(Val(parse(1))).Tile(X, y).M2Anim = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).M2AnimSet = Val(parse(n + 1))
            Case 5
                Map(Val(parse(1))).Tile(X, y).Fringe = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).FringeSet = Val(parse(n + 1))
            Case 6
                Map(Val(parse(1))).Tile(X, y).FAnim = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).FAnimSet = Val(parse(n + 1))
            Case 7
                Map(Val(parse(1))).Tile(X, y).Fringe2 = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).Fringe2Set = Val(parse(n + 1))
            Case 8
                Map(Val(parse(1))).Tile(X, y).F2Anim = Val(parse(n))
                Map(Val(parse(1))).Tile(X, y).F2AnimSet = Val(parse(n + 1))
        End Select
        Call SaveLocalMap(Val(parse(1)))
    End If

    If casestring = "tilecheckattribute" Then
        n = 5
        X = Val(parse(2))
        y = Val(parse(3))

        Map(Val(parse(1))).Tile(X, y).Type = Val(parse(n - 1))
        Map(Val(parse(1))).Tile(X, y).Data1 = Val(parse(n))
        Map(Val(parse(1))).Tile(X, y).Data2 = Val(parse(n + 1))
        Map(Val(parse(1))).Tile(X, y).Data3 = Val(parse(n + 2))
        Map(Val(parse(1))).Tile(X, y).String1 = parse(n + 3)
        Map(Val(parse(1))).Tile(X, y).String2 = parse(n + 4)
        Map(Val(parse(1))).Tile(X, y).String3 = parse(n + 5)
        Call SaveLocalMap(Val(parse(1)))
    End If

    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If casestring = "mapitemdata" Then
        n = 1

        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).num = Val(parse(n))
            SaveMapItem(i).Value = Val(parse(n + 1))
            SaveMapItem(i).Dur = Val(parse(n + 2))
            SaveMapItem(i).X = Val(parse(n + 3))
            SaveMapItem(i).y = Val(parse(n + 4))

            n = n + 5
        Next i

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If casestring = "mapnpcdata" Then
        n = 1

        For i = 1 To 15
            SaveMapNpc(i).num = Val(parse(n))
            SaveMapNpc(i).X = Val(parse(n + 1))
            SaveMapNpc(i).y = Val(parse(n + 2))
            SaveMapNpc(i).Dir = Val(parse(n + 3))

            n = n + 4
        Next i

        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If casestring = "mapdone" Then
        ' Map = SaveMap

        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i

        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i

        GettingMap = False

        ' Play music
        If Trim$(Map(GetPlayerMap(MyIndex)).music) <> "None" Then
            Call MapMusic(Map(GetPlayerMap(MyIndex)).music)
        End If

        If GameWeather = WEATHER_RAINING Then
            Call PlayBGS("rain.wav")
        End If
        If GameWeather = WEATHER_THUNDER Then
            Call PlayBGS("thunder.wav")
        End If

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (casestring = "saymsg") Or (casestring = "broadcastmsg") Or (casestring = "globalmsg") Or (casestring = "playermsg") Or (casestring = "mapmsg") Or (casestring = "adminmsg") Then
        Call AddText(parse(1), Val(parse(2)))
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If casestring = "spawnitem" Then
        n = Val(parse(1))

        MapItem(n).num = Val(parse(2))
        MapItem(n).Value = Val(parse(3))
        MapItem(n).Dur = Val(parse(4))
        MapItem(n).X = Val(parse(5))
        MapItem(n).y = Val(parse(6))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "itemeditor") Then
        InItemsEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Item(i).Name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updateitem") Then
        n = Val(parse(1))

        ' Update the item
        Item(n).Name = parse(2)
        Item(n).Pic = Val(parse(3))
        Item(n).Type = Val(parse(4))
        Item(n).Data1 = Val(parse(5))
        Item(n).Data2 = Val(parse(6))
        Item(n).Data3 = Val(parse(7))
        Item(n).StrReq = Val(parse(8))
        Item(n).DefReq = Val(parse(9))
        Item(n).SpeedReq = Val(parse(10))
        Item(n).MagicReq = Val(parse(11))
        Item(n).ClassReq = Val(parse(12))
        Item(n).AccessReq = Val(parse(13))

        Item(n).AddHP = Val(parse(14))
        Item(n).AddMP = Val(parse(15))
        Item(n).AddSP = Val(parse(16))
        Item(n).AddSTR = Val(parse(17))
        Item(n).AddDEF = Val(parse(18))
        Item(n).AddMAGI = Val(parse(19))
        Item(n).AddSpeed = Val(parse(20))
        Item(n).AddEXP = Val(parse(21))
        Item(n).desc = parse(22)
        Item(n).AttackSpeed = Val(parse(23))
        Item(n).Price = Val(parse(24))
        Item(n).Stackable = Val(parse(25))
        Item(n).Bound = Val(parse(26))
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (casestring = "edititem") Then
        n = Val(parse(1))

        ' Update the item
        Item(n).Name = parse(2)
        Item(n).Pic = Val(parse(3))
        Item(n).Type = Val(parse(4))
        Item(n).Data1 = Val(parse(5))
        Item(n).Data2 = Val(parse(6))
        Item(n).Data3 = Val(parse(7))
        Item(n).StrReq = Val(parse(8))
        Item(n).DefReq = Val(parse(9))
        Item(n).SpeedReq = Val(parse(10))
        Item(n).MagicReq = Val(parse(11))
        Item(n).ClassReq = Val(parse(12))
        Item(n).AccessReq = Val(parse(13))

        Item(n).AddHP = Val(parse(14))
        Item(n).AddMP = Val(parse(15))
        Item(n).AddSP = Val(parse(16))
        Item(n).AddSTR = Val(parse(17))
        Item(n).AddDEF = Val(parse(18))
        Item(n).AddMAGI = Val(parse(19))
        Item(n).AddSpeed = Val(parse(20))
        Item(n).AddEXP = Val(parse(21))
        Item(n).desc = parse(22)
        Item(n).AttackSpeed = Val(parse(23))
        Item(n).Price = Val(parse(24))
        Item(n).Stackable = Val(parse(25))
        Item(n).Bound = Val(parse(26))

        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: mouse packet  ::
    ' :::::::::::::::::::
    If (casestring = "mouse") Then
        Player(MyIndex).input = 1
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' ::Weather Packet::
    ' ::::::::::::::::::
    If (casestring = "mapweather") Then
        If 0 + Val(parse(1)) <> 0 Then
            Map(Val(parse(1))).Weather = Val(parse(2))
            If Val(parse(1)) = 2 Then
                frmMirage.tmrSnowDrop.Interval = Val(parse(3))
            ElseIf Val(parse(1)) = 1 Then
                frmMirage.tmrRainDrop.Interval = Val(parse(3))
            End If
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If casestring = "spawnnpc" Then
        n = Val(parse(1))

        MapNpc(n).num = Val(parse(2))
        MapNpc(n).X = Val(parse(3))
        MapNpc(n).y = Val(parse(4))
        MapNpc(n).Dir = Val(parse(5))
        MapNpc(n).Big = Val(parse(6))

        ' Client use only
        MapNpc(n).xOffset = 0
        MapNpc(n).yOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If casestring = "npcdead" Then
        n = Val(parse(1))

        MapNpc(n).num = 0
        MapNpc(n).X = 0
        MapNpc(n).y = 0
        MapNpc(n).Dir = 0

        ' Client use only
        MapNpc(n).xOffset = 0
        MapNpc(n).yOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (casestring = "npceditor") Then
        InNpcEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Npc(i).Name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (casestring = "updatenpc") Then
        n = Val(parse(1))

        ' Update the item
        Npc(n).Name = parse(2)
        Npc(n).AttackSay = vbNullString
        Npc(n).Sprite = Val(parse(3))
        Npc(n).SpriteSize = Val(parse(4))
        
        ' That's all well and good but it also resets our NPC - Pickle
        ' Npc(n).SpawnSecs = 0
        ' Npc(n).Behavior = 0
        ' Npc(n).Range = 0
        ' For i = 1 To MAX_NPC_DROPS
        ' Npc(n).ItemNPC(i).chance = 0
        ' Npc(n).ItemNPC(i).ItemNum = 0
        ' Npc(n).ItemNPC(i).ItemValue = 0
        ' Next i
        ' Npc(n).STR = 0
        ' Npc(n).DEF = 0
        ' Npc(n).speed = 0
        ' Npc(n).MAGI = 0
        
        Npc(n).Big = Val(parse(5))
        Npc(n).MaxHp = Val(parse(6))
        ' Npc(n).Exp = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    If (casestring = "editnpc") Then
        n = Val(parse(1))

        ' Update the npc
        Npc(n).Name = parse(2)
        Npc(n).AttackSay = parse(3)
        Npc(n).Sprite = Val(parse(4))
        Npc(n).SpawnSecs = Val(parse(5))
        Npc(n).Behavior = Val(parse(6))
        Npc(n).Range = Val(parse(7))
        Npc(n).STR = Val(parse(8))
        Npc(n).DEF = Val(parse(9))
        Npc(n).speed = Val(parse(10))
        Npc(n).MAGI = Val(parse(11))
        Npc(n).Big = Val(parse(12))
        Npc(n).MaxHp = Val(parse(13))
        Npc(n).Exp = Val(parse(14))
        Npc(n).SpawnTime = Val(parse(15))
        Npc(n).Element = Val(parse(16))
        Npc(n).SpriteSize = Val(parse(17))

        ' Call GlobalMsg("At editnpc..." & Npc(n).Element)
        z = 18
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).chance = Val(parse(z))
            Npc(n).ItemNPC(i).ItemNum = Val(parse(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(parse(z + 2))
            z = z + 3
        Next i

        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (casestring = "mapkey") Then
        X = Val(parse(1))
        y = Val(parse(2))
        n = Val(parse(3))

        TempTile(X, y).DoorOpen = n

        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (casestring = "editmap") Then
        Call EditorInit
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "shopeditor") Then
        InShopEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Shop(i).Name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updateshop") Then
        n = Val(parse(1))

        ' Update the shop name
        Shop(n).Name = parse(2)
        Shop(n).FixesItems = Val(parse(3))
        Shop(n).BuysItems = Val(parse(4))
        Shop(n).ShowInfo = Val(parse(5))
        Shop(n).currencyItem = Val(parse(6))

        m = 7
        ' Get shop items
        For i = 1 To MAX_SHOP_ITEMS
            Shop(n).ShopItem(i).ItemNum = Val(parse(m))
            Shop(n).ShopItem(i).Amount = Val(parse(m + 1))
            Shop(n).ShopItem(i).Price = Val(parse(m + 2))
            m = m + 3
        Next i

        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (casestring = "editshop") Then

        shopNum = Val(parse(1))

        ' Update the shop
        Shop(shopNum).Name = parse(2)
        Shop(shopNum).FixesItems = Val(parse(3))
        Shop(shopNum).BuysItems = Val(parse(4))
        Shop(shopNum).ShowInfo = Val(parse(5))
        Shop(shopNum).currencyItem = Val(parse(6))

        m = 7
        For i = 1 To 25
            Shop(shopNum).ShopItem(i).ItemNum = Val(parse(m))
            Shop(shopNum).ShopItem(i).Amount = Val(parse(m + 1))
            Shop(shopNum).ShopItem(i).Price = Val(parse(m + 2))
            m = m + 3
        Next i

        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (casestring = "spelleditor") Then
        InSpellEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Spell(i).Name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (casestring = "updatespell") Then
        n = Val(parse(1))

        ' Update the spell name
        Spell(n).Name = parse(2)
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (casestring = "editspell") Then
        n = Val(parse(1))

        ' Update the spell
        Spell(n).Name = parse(2)
        Spell(n).ClassReq = Val(parse(3))
        Spell(n).LevelReq = Val(parse(4))
        Spell(n).Type = Val(parse(5))
        Spell(n).Data1 = Val(parse(6))
        Spell(n).Data2 = Val(parse(7))
        Spell(n).Data3 = Val(parse(8))
        Spell(n).MPCost = Val(parse(9))
        Spell(n).Sound = Val(parse(10))
        Spell(n).Range = Val(parse(11))
        Spell(n).SpellAnim = Val(parse(12))
        Spell(n).SpellTime = Val(parse(13))
        Spell(n).SpellDone = Val(parse(14))
        Spell(n).AE = Val(parse(15))
        Spell(n).Big = Val(parse(16))
        Spell(n).Element = Val(parse(17))


        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (casestring = "goshop") Then
        shopNum = Val(parse(1))
        ' Show the shop
        Call GoShop(shopNum)
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (casestring = "spells") Then

        frmMirage.picPlayerSpells.Visible = True
        frmMirage.lstSpells.Clear

        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(parse(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmMirage.lstSpells.addItem i & ": " & Trim$(Spell(Player(MyIndex).Spell(i)).Name)
            Else
                frmMirage.lstSpells.addItem "--- Slot Free ---"
            End If
        Next i

        frmMirage.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (casestring = "weather") Then
        If Val(parse(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
            Call AddText("You see drops of rain falling from the sky above!", BRIGHTGREEN)
            Call PlayBGS("rain.mp3")
        End If
        If Val(parse(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
            Call AddText("You see thunder in the sky above!", BRIGHTGREEN)
            Call PlayBGS("thunder.mp3")
        End If
        If Val(parse(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("You see snow falling from the sky above!", BRIGHTGREEN)
        End If

        If Val(parse(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("The rain beings to calm.", BRIGHTGREEN)
                Call frmMirage.BGSPlayer.StopMedia
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("The snow is melting away.", BRIGHTGREEN)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("The thunder begins to disapear.", BRIGHTGREEN)
                Call frmMirage.BGSPlayer.StopMedia
            End If
        End If
        GameWeather = Val(parse(1))
        RainIntensity = Val(parse(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If

    ' ::::::::::::::::::::::::::::::::
    ' :: playername coloring packet ::
    ' ::::::::::::::::::::::::::::::::
    If (casestring = "namecolor") Then
        Player(MyIndex).color = Val(parse(1))
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: image packet      ::
    ' :::::::::::::::::::::::
    If (LCase$(parse(0)) = "fog") Then
        rec.Top = Int(Val(parse(4)))
        rec.Bottom = Int(Val(parse(5)))
        rec.Left = Int(Val(parse(6)))
        rec.Right = Int(Val(parse(7)))
        Call DD_BackBuffer.BltFast(Val(parse(1)), Val(parse(2)), DD_TileSurf(Val(parse(3))), rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If casestring = "onlinelist" Then
        frmMirage.lstOnline.Clear

        n = 2
        z = Val(parse(1))
        For X = n To (z + 1)
            frmMirage.lstOnline.addItem Trim$(parse(n))
            n = n + 2
        Next X
        Exit Sub
    End If

    ' ::::::::::::::::::::::::
    ' :: Blit Player Damage ::
    ' ::::::::::::::::::::::::
    If casestring = "blitplayerdmg" Then
        DmgDamage = Val(parse(1))
        NPCWho = Val(parse(2))
        DmgTime = GetTickCount
        iii = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Blit NPC Damage ::
    ' :::::::::::::::::::::
    If casestring = "blitnpcdmg" Then
        NPCDmgDamage = Val(parse(1))
        NPCDmgTime = GetTickCount
        ii = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If casestring = "pptrading" Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
            frmPlayerTrade.Items1.addItem i & ": <Nothing>"
            frmPlayerTrade.Items2.addItem i & ": <Nothing>"
        Next i

        frmPlayerTrade.Items1.ListIndex = 0

        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "qtrade" Then
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = vbNullString
            Trading2(i).InvNum = 0
            Trading2(i).InvName = vbNullString
        Next i

        frmPlayerTrade.Command1.ForeColor = &H0&
        frmPlayerTrade.Command2.ForeColor = &H0&

        frmPlayerTrade.Visible = False
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Disable Time ::
    ' ::::::::::::::::::
    If casestring = "dtime" Then
        If parse(1) = "True" Then
            frmMirage.lblGameTime.Caption = vbNullString
            frmMirage.lblGameClock.Caption = vbNullString
            frmMirage.lblGameTime.Visible = False
            frmMirage.lblGameClock.Visible = False
            frmMirage.tmrGameClock.Enabled = False
        Else
            frmMirage.lblGameTime.Caption = "It is now:"
            frmMirage.lblGameClock.Caption = vbNullString
            frmMirage.lblGameTime.Visible = True
            frmMirage.lblGameClock.Visible = True
            frmMirage.tmrGameClock.Enabled = True
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "updatetradeitem" Then
        n = Val(parse(1))

        Trading2(n).InvNum = Val(parse(2))
        Trading2(n).InvName = parse(3)

        If STR(Trading2(n).InvNum) <= 0 Then
            frmPlayerTrade.Items2.List(n - 1) = n & ": <Nothing>"
        Else
            frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim$(Trading2(n).InvName)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "trading" Then
        n = Val(parse(1))
        If n = 0 Then
            frmPlayerTrade.Command2.ForeColor = &H0&
        End If
        If n = 1 Then
            frmPlayerTrade.Command2.ForeColor = &HFF00&
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Chat System Packets ::
    ' :::::::::::::::::::::::::
    If casestring = "ppchatting" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Label1.Caption = "Chatting With: " & Trim$(Player(Val(parse(1))).Name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If

    If casestring = "qchat" Then
        frmPlayerChat.txtChat.Text = vbNullString
        frmPlayerChat.txtSay.Text = vbNullString
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    End If

    If casestring = "sendchat" Then
        Dim s As String

        s = vbNewLine & GetPlayerName(Val(parse(2))) & "> " & Trim$(parse(1))
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
        frmPlayerChat.txtChat.SelColor = QBColor(BROWN)
        frmPlayerChat.txtChat.SelText = s
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text) - 1
        Exit Sub
    End If
' :::::::::::::::::::::::::::::
' :: END Chat System Packets ::
' :::::::::::::::::::::::::::::

    ' :::::::::::::::::::::::
    ' :: Play Sound Packet ::
    ' :::::::::::::::::::::::
    If casestring = "sound" Then
        s = LCase$(parse(1))
        Select Case Trim$(s)
            Case "attack"
                Call PlaySound("sword.wav")
            Case "critical"
                Call PlaySound("critical.wav")
            Case "miss"
                Call PlaySound("miss.wav")
            Case "key"
                Call PlaySound("key.wav")
            Case "magic"
                Call PlaySound("magic" & Val(parse(2)) & ".wav")
            Case "warp"
                If FileExists("SFX\warp.wav") Then
                    Call PlaySound("warp.wav")
                End If
            Case "pain"
                Call PlaySound("pain.wav")
            Case "soundattribute"
                Call PlaySound(Trim$(parse(2)))
        End Select
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "spritechange" Then
        If Val(parse(1)) = 1 Then
            i = MsgBox("Would you like to buy this sprite?", 4, "Buying Sprite")
            If i = 6 Then
                Call SendData("buysprite" & END_CHAR)
            End If
        Else
            Call SendData("buysprite" & END_CHAR)
        End If
        Exit Sub
    End If
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: House Buy Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If casestring = "housebuy" Then
        If Val(parse(1)) = 1 Then
            i = MsgBox("Would you like to buy this house?", 4, "Buying House")
            If i = 6 Then
                Call SendData("buyhouse" & END_CHAR)
            End If
        Else
            Call SendData("buyhouse" & END_CHAR)
        End If
        Exit Sub
    End If
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If casestring = "changedir" Then
        Player(Val(parse(2))).Dir = Val(parse(1))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::
    ' :: Flash Movie Event Packet ::
    ' ::::::::::::::::::::::::::::::
    If casestring = "flashevent" Then
        If LCase$(Mid$(Trim$(parse(1)), 1, 7)) = "http://" Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, Trim$(parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        ElseIf FileExists("Flashs\" & Trim$(parse(1))) = True Then
            WriteINI "CONFIG", "Music", 0, App.Path & "\config.ini"
            WriteINI "CONFIG", "Sound", 0, App.Path & "\config.ini"
            Call StopBGM
            Call StopSound
            frmFlash.Flash.LoadMovie 0, App.Path & "\Flashs\" & Trim$(parse(1))
            frmFlash.Flash.Play
            frmFlash.Check.Enabled = True
            frmFlash.Show vbModeless, frmMirage
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If casestring = "prompt" Then
        i = MsgBox(Trim$(parse(1)), vbYesNo)
        Call SendData("prompt" & SEP_CHAR & i & SEP_CHAR & Val(parse(2)) & END_CHAR)
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If casestring = "querybox" Then
        frmQuery.Label1.Caption = Trim$(parse(1))
        frmQuery.Label2.Caption = parse(2)
        frmQuery.Show vbModal
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "emoticoneditor") Then
        InEmoticonEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For i = 0 To MAX_EMOTICONS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Emoticons(i).Command)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Element editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "elementeditor") Then
        InElementEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For i = 0 To MAX_ELEMENTS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Element(i).Name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = "editelement") Then
        n = Val(parse(1))

        Element(n).Name = parse(2)
        Element(n).Strong = Val(parse(3))
        Element(n).Weak = Val(parse(4))

        Call ElementEditorInit
        Exit Sub
    End If

    If (casestring = "updateelement") Then
        n = Val(parse(1))

        Element(n).Name = parse(2)
        Element(n).Strong = Val(parse(3))
        Element(n).Weak = Val(parse(4))
        Exit Sub
    End If

    If (casestring = "editemoticon") Then
        n = Val(parse(1))

        Emoticons(n).Command = parse(2)
        Emoticons(n).Pic = Val(parse(3))

        Call EmoticonEditorInit
        Exit Sub
    End If

    If (casestring = "updateemoticon") Then
        n = Val(parse(1))

        Emoticons(n).Command = parse(2)
        Emoticons(n).Pic = Val(parse(3))
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Arrow editor packet ::
    ' ::::::::::::::::::::::::::::
    If (casestring = "arroweditor") Then
        InArrowEditor = True

        frmIndex.Show
        frmIndex.lstIndex.Clear

        For i = 1 To MAX_ARROWS
            frmIndex.lstIndex.addItem i & ": " & Trim$(Arrows(i).Name)
        Next i

        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If

    If (casestring = "updatearrow") Then
        n = Val(parse(1))

        Arrows(n).Name = parse(2)
        Arrows(n).Pic = Val(parse(3))
        Arrows(n).Range = Val(parse(4))
        Arrows(n).Amount = Val(parse(5))
        Exit Sub
    End If

    If (casestring = "editarrow") Then
        n = Val(parse(1))

        Arrows(n).Name = parse(2)

        Call ArrowEditorInit
        Exit Sub
    End If

    If (casestring = "updatearrow") Then
        n = Val(parse(1))

        Arrows(n).Name = parse(2)
        Arrows(n).Pic = Val(parse(3))
        Arrows(n).Range = Val(parse(4))
        Arrows(n).Amount = Val(parse(5))
        Exit Sub
    End If

    If (casestring = "hookshot") Then
        n = Val(parse(1))
        i = Val(parse(3))

        Player(n).HookShotAnim = Arrows(Val(parse(2))).Pic
        Player(n).HookShotTime = GetTickCount
        Player(n).HookShotToX = Val(parse(4))
        Player(n).HookShotToY = Val(parse(5))
        Player(n).HookShotX = GetPlayerX(n)
        Player(n).HookShotY = GetPlayerY(n)
        Player(n).HookShotSucces = Val(parse(6))
        Player(n).HookShotDir = Val(parse(3))

        Call PlaySound("grapple.wav")
        Call PlaySound("grapple-fire.wav")

        If i = DIR_DOWN Then
            Player(n).HookShotX = GetPlayerX(n)
            Player(n).HookShotY = GetPlayerY(n) + 1
            If Player(n).HookShotX - 1 > MAX_MAPY Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If i = DIR_UP Then
            Player(n).HookShotX = GetPlayerX(n)
            Player(n).HookShotY = GetPlayerY(n) - 1
            If Player(n).HookShotY + 1 < 0 Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If i = DIR_RIGHT Then
            Player(n).HookShotX = GetPlayerX(n) + 1
            Player(n).HookShotY = GetPlayerY(n)
            If Player(n).HookShotX - 1 > MAX_MAPX Then
                Player(n).HookShotX = 0
                Player(n).HookShotY = 0
                Exit Sub
            End If
        End If
        If i = DIR_LEFT Then
            Player(n).HookShotX = GetPlayerX(n) - 1
            Player(n).HookShotY = GetPlayerY(n)
            If Player(n).HookShotX + 1 < 0 Then
                Player(n).Arrow(X).Arrow = 0
                Exit Sub
            End If
        End If
        Exit Sub
    End If

    If (casestring = "checkarrows") Then
        n = Val(parse(1))
        z = Val(parse(2))
        i = Val(parse(3))

        For X = 1 To MAX_PLAYER_ARROWS
            If Player(n).Arrow(X).Arrow = 0 Then
                Player(n).Arrow(X).Arrow = 1
                Player(n).Arrow(X).ArrowNum = z
                Player(n).Arrow(X).ArrowAnim = Arrows(z).Pic
                Player(n).Arrow(X).ArrowTime = GetTickCount
                Player(n).Arrow(X).ArrowVarX = 0
                Player(n).Arrow(X).ArrowVarY = 0
                Player(n).Arrow(X).ArrowY = GetPlayerY(n)
                Player(n).Arrow(X).ArrowX = GetPlayerX(n)
                Player(n).Arrow(X).ArrowAmount = p

                If i = DIR_DOWN Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 0
                    If Player(n).Arrow(X).ArrowY - 1 > MAX_MAPY Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_UP Then
                    Player(n).Arrow(X).ArrowY = GetPlayerY(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 1
                    If Player(n).Arrow(X).ArrowY + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_RIGHT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) + 1
                    Player(n).Arrow(X).ArrowPosition = 2
                    If Player(n).Arrow(X).ArrowX - 1 > MAX_MAPX Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                If i = DIR_LEFT Then
                    Player(n).Arrow(X).ArrowX = GetPlayerX(n) - 1
                    Player(n).Arrow(X).ArrowPosition = 3
                    If Player(n).Arrow(X).ArrowX + 1 < 0 Then
                        Player(n).Arrow(X).Arrow = 0
                        Exit Sub
                    End If
                End If
                Exit For
            End If
        Next X
        Exit Sub
    End If

    If (casestring = "checksprite") Then
        n = Val(parse(1))

        Player(n).Sprite = Val(parse(2))
        Exit Sub
    End If

    If (casestring = "mapreport") Then
        n = 1

        frmMapReport.lstIndex.Clear
        For i = 1 To MAX_MAPS
            frmMapReport.lstIndex.addItem i & ": " & Trim$(parse(n))
            n = n + 1
        Next i

        frmMapReport.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (casestring = "time") Then
        GameTime = Val(parse(1))
        If GameTime = TIME_DAY Then
            Call AddText("Day has dawned in this realm.", WHITE)
        Else
            Call AddText("Night has fallen upon the weary eyed nightowls.", WHITE)
        End If
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = "spellanim") Then
        Dim SpellNum As Long
        SpellNum = Val(parse(1))

        Spell(SpellNum).SpellAnim = Val(parse(2))
        Spell(SpellNum).SpellTime = Val(parse(3))
        Spell(SpellNum).SpellDone = Val(parse(4))
        Spell(SpellNum).Big = Val(parse(9))

        Player(Val(parse(5))).SpellNum = SpellNum

        For i = 1 To MAX_SPELL_ANIM
            If Player(Val(parse(5))).SpellAnim(i).CastedSpell = NO Then
                Player(Val(parse(5))).SpellAnim(i).SpellDone = 0
                Player(Val(parse(5))).SpellAnim(i).SpellVar = 0
                Player(Val(parse(5))).SpellAnim(i).SpellTime = GetTickCount
                Player(Val(parse(5))).SpellAnim(i).TargetType = Val(parse(6))
                Player(Val(parse(5))).SpellAnim(i).Target = Val(parse(7))
                Player(Val(parse(5))).SpellAnim(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If

    ' :::::::::::::::::::::::
    ' :: Script Spell anim packet ::
    ' :::::::::::::::::::::::
    If (casestring = "scriptspellanim") Then
        Spell(Val(parse(1))).SpellAnim = Val(parse(2))
        Spell(Val(parse(1))).SpellTime = Val(parse(3))
        Spell(Val(parse(1))).SpellDone = Val(parse(4))
        Spell(Val(parse(1))).Big = Val(parse(7))


        For i = 1 To MAX_SCRIPTSPELLS
            If ScriptSpell(i).CastedSpell = NO Then
                ScriptSpell(i).SpellNum = Val(parse(1))
                ScriptSpell(i).SpellDone = 0
                ScriptSpell(i).SpellVar = 0
                ScriptSpell(i).SpellTime = GetTickCount
                ScriptSpell(i).X = Val(parse(5))
                ScriptSpell(i).y = Val(parse(6))
                ScriptSpell(i).CastedSpell = YES
                Exit For
            End If
        Next i
        Exit Sub
    End If

    If (casestring = "checkemoticons") Then
        n = Val(parse(1))

        Player(n).EmoticonNum = Val(parse(2))
        Player(n).EmoticonTime = GetTickCount
        Player(n).EmoticonVar = 0
        Exit Sub
    End If


    If casestring = "levelup" Then
        Player(Val(parse(1))).LevelUpT = GetTickCount
        Player(Val(parse(1))).LevelUp = 1
        Exit Sub
    End If

    If casestring = "damagedisplay" Then
        For i = 1 To MAX_BLT_LINE
            If Val(parse(1)) = 0 Then
                If BattlePMsg(i).Index <= 0 Then
                    BattlePMsg(i).Index = 1
                    BattlePMsg(i).Msg = parse(2)
                    BattlePMsg(i).color = Val(parse(3))
                    BattlePMsg(i).Time = GetTickCount
                    BattlePMsg(i).Done = 1
                    BattlePMsg(i).y = 0
                    Exit Sub
                Else
                    BattlePMsg(i).y = BattlePMsg(i).y - 15
                End If
            Else
                If BattleMMsg(i).Index <= 0 Then
                    BattleMMsg(i).Index = 1
                    BattleMMsg(i).Msg = parse(2)
                    BattleMMsg(i).color = Val(parse(3))
                    BattleMMsg(i).Time = GetTickCount
                    BattleMMsg(i).Done = 1
                    BattleMMsg(i).y = 0
                    Exit Sub
                Else
                    BattleMMsg(i).y = BattleMMsg(i).y - 15
                End If
            End If
        Next i

        z = 1
        If Val(parse(1)) = 0 Then
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattlePMsg(i).y < BattlePMsg(i + 1).y Then
                        z = i
                    End If
                Else
                    If BattlePMsg(i).y < BattlePMsg(1).y Then
                        z = i
                    End If
                End If
            Next i

            BattlePMsg(z).Index = 1
            BattlePMsg(z).Msg = parse(2)
            BattlePMsg(z).color = Val(parse(3))
            BattlePMsg(z).Time = GetTickCount
            BattlePMsg(z).Done = 1
            BattlePMsg(z).y = 0
        Else
            For i = 1 To MAX_BLT_LINE
                If i < MAX_BLT_LINE Then
                    If BattleMMsg(i).y < BattleMMsg(i + 1).y Then
                        z = i
                    End If
                Else
                    If BattleMMsg(i).y < BattleMMsg(1).y Then
                        z = i
                    End If
                End If
            Next i

            BattleMMsg(z).Index = 1
            BattleMMsg(z).Msg = parse(2)
            BattleMMsg(z).color = Val(parse(3))
            BattleMMsg(z).Time = GetTickCount
            BattleMMsg(z).Done = 1
            BattleMMsg(z).y = 0
        End If
        Exit Sub
    End If

    If casestring = "itembreak" Then
        ItemDur(Val(parse(1))).Item = Val(parse(2))
        ItemDur(Val(parse(1))).Dur = Val(parse(3))
        ItemDur(Val(parse(1))).Done = 1
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::::::::::::::
    ' :: Index player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::::::::
    If casestring = "itemworn" Then
        Player(Val(parse(1))).Armor = Val(parse(2))
        Player(Val(parse(1))).Weapon = Val(parse(3))
        Player(Val(parse(1))).Helmet = Val(parse(4))
        Player(Val(parse(1))).Shield = Val(parse(5))
        Player(Val(parse(1))).legs = Val(parse(6))
        Player(Val(parse(1))).Ring = Val(parse(7))
        Player(Val(parse(1))).Necklace = Val(parse(8))
        Exit Sub
    End If

    If casestring = "scripttile" Then
        frmScript.lblScript.Caption = parse(1)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Set player speed ::
    ' ::::::::::::::::::::::
    If casestring = "setspeed" Then
        SetSpeed parse(1), Val(parse(2))
        Exit Sub
    End If

    ' ::::::::::::::::::
    ' :: Custom Menu  ::
    ' ::::::::::::::::::
    If (casestring = "showcustommenu") Then
        ' Error handling
        If Not FileExists(parse(2)) Then
            Call MsgBox(parse(2) & " not found. Menu loading aborted. Please contact a GM to fix this problem.", vbExclamation)
            Exit Sub
        End If

        CUSTOM_TITLE = parse(1)
        CUSTOM_IS_CLOSABLE = Val(parse(3))

        frmCustom1.picBackground.Top = 0
        frmCustom1.picBackground.Left = 0
        frmCustom1.picBackground = LoadPicture(App.Path & parse(2))
        frmCustom1.Height = PixelsToTwips(24 + frmCustom1.picBackground.Height, 1)
        frmCustom1.Width = PixelsToTwips(6 + frmCustom1.picBackground.Width, 0)
        frmCustom1.Visible = True

        Exit Sub
    End If

    If (casestring = "closecustommenu") Then

        CUSTOM_TITLE = "CLOSED"
        Unload frmCustom1

        Exit Sub
    End If

    If (casestring = "loadpiccustommenu") Then

        CustomIndex = parse(1)
        strfilename = parse(2)
        CustomX = Val(parse(3))
        CustomY = Val(parse(4))
        
        If CustomIndex > frmCustom1.picCustom.UBound Then
            Load frmCustom1.picCustom(CustomIndex)
        End If

        If strfilename = vbNullString Then
            strfilename = "MEGAUBERBLANKNESSOFUNHOLYPOWER" 'smooth :\    -Pickle
        End If

        If FileExists(strfilename) = True Then
            frmCustom1.picCustom(CustomIndex) = LoadPicture(App.Path & strfilename)
            frmCustom1.picCustom(CustomIndex).Top = CustomY
            frmCustom1.picCustom(CustomIndex).Left = CustomX
            frmCustom1.picCustom(CustomIndex).Visible = True
        Else
            frmCustom1.picCustom(CustomIndex).Picture = LoadPicture()
            frmCustom1.picCustom(CustomIndex).Visible = False
        End If

        Exit Sub
    End If

    If (casestring = "loadlabelcustommenu") Then

        CustomIndex = parse(1)
        strfilename = parse(2)
        CustomX = Val(parse(3))
        CustomY = Val(parse(4))
        customsize = Val(parse(5))
        customcolour = Val(parse(6))
        
        If CustomIndex > frmCustom1.BtnCustom.UBound Then
            Load frmCustom1.BtnCustom(CustomIndex)
        End If

        frmCustom1.BtnCustom(CustomIndex).Caption = strfilename
        frmCustom1.BtnCustom(CustomIndex).Top = CustomY
        frmCustom1.BtnCustom(CustomIndex).Left = CustomX
        frmCustom1.BtnCustom(CustomIndex).Font.Bold = True
        frmCustom1.BtnCustom(CustomIndex).Font.Size = customsize
        frmCustom1.BtnCustom(CustomIndex).ForeColor = QBColor(customcolour)
        frmCustom1.BtnCustom(CustomIndex).Visible = True
        frmCustom1.BtnCustom(CustomIndex).Alignment = parse(7)

        If parse(8) <= 0 Or parse(9) <= 0 Then
            frmCustom1.BtnCustom(CustomIndex).AutoSize = True
        Else
            frmCustom1.BtnCustom(CustomIndex).AutoSize = False
            frmCustom1.BtnCustom(CustomIndex).Width = parse(8)
            frmCustom1.BtnCustom(CustomIndex).Height = parse(9)
        End If

        Exit Sub
    End If

    If (casestring = "loadtextboxcustommenu") Then

        CustomIndex = parse(1)
        strfilename = parse(2)
        CustomX = Val(parse(3))
        CustomY = Val(parse(4))
        customtext = parse(5)
        
        If CustomIndex > frmCustom1.txtCustom.UBound Then
            Load frmCustom1.txtCustom(CustomIndex)
            Load frmCustom1.txtcustomOK(CustomIndex)
        End If

        frmCustom1.txtCustom(CustomIndex).Text = customtext
        frmCustom1.txtCustom(CustomIndex).Top = CustomY
        frmCustom1.txtCustom(CustomIndex).Left = strfilename
        frmCustom1.txtCustom(CustomIndex).Width = CustomX - 32
        frmCustom1.txtcustomOK(CustomIndex).Top = CustomY
        frmCustom1.txtcustomOK(CustomIndex).Left = frmCustom1.txtCustom(CustomIndex).Left + frmCustom1.txtCustom(CustomIndex).Width
        frmCustom1.txtcustomOK(CustomIndex).Visible = True
        frmCustom1.txtCustom(CustomIndex).Visible = True

        Exit Sub
    End If

    If (casestring = "loadinternetwindow") Then
        customtext = parse(1)
        ' DEBUG STRING
        ' Call AddText(customtext, 15)
        ShellExecute 1, "open", Trim(customtext), vbNullString, vbNullString, 1
        Exit Sub
    End If

    If (casestring = "returncustomboxmsg") Then
        customsize = parse(1)

        packet = "returningcustomboxmsg" & SEP_CHAR & frmCustom1.txtCustom(customsize).Text & END_CHAR
        Call SendData(packet)

        Exit Sub
    End If

End Sub
