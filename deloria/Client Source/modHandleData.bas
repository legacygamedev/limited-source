Attribute VB_Name = "modHandleData"
Sub HandleData(ByVal Data As String)
Dim Parse() As String
Dim Name As String
Dim Password As String
Dim Sex As Long
Dim ClassNum As Long
Dim CharNum As Long
Dim Msg As String
Dim IPMask As String
Dim BanSlot As Long
Dim MsgTo As Long
Dim Dir As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim PointType As Long
Dim BanPlayer As Long
Dim Level As Long
Dim i As Long, n As Long, x As Long, Y As Long
Dim ShopNum As Long, GiveItem As Long, GiveValue As Long, GetItem As Long, GetValue As Long
Dim z As Long

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If DebugMode = True Then
        If frmDebug.Visible = False Then frmDebug.Visible = True
        Call TextAdd(frmDebug.txtDebug, "((( Processed Packet " & Parse(0) & " )))", True)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Get players stats ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "maxinfo" Then
        GAME_NAME = Trim(Parse(1))
        MAX_PLAYERS = Val(Parse(2))
        MAX_ITEMS = Val(Parse(3))
        MAX_NPCS = Val(Parse(4))
        MAX_SHOPS = Val(Parse(5))
        MAX_SPELLS = Val(Parse(6))
        MAX_MAPS = Val(Parse(7))
        MAX_MAP_ITEMS = Val(Parse(8))
        MAX_MAPX = Val(Parse(9))
        MAX_MAPY = Val(Parse(10))
        MAX_EMOTICONS = Val(Parse(11))

        ReDim Player(1 To MAX_PLAYERS) As PlayerRec
        ReDim Item(1 To MAX_ITEMS) As ItemRec
        ReDim Npc(1 To MAX_NPCS) As NpcRec
        ReDim MapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim Shop(1 To MAX_SHOPS) As ShopRec
        ReDim Spell(1 To MAX_SPELLS) As SpellRec
        ReDim Bubble(1 To MAX_PLAYERS) As ChatBubble
        ReDim SaveMapItem(1 To MAX_MAP_ITEMS) As MapItemRec
        ReDim SaveMap.Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        ReDim TempTile(0 To MAX_MAPX, 0 To MAX_MAPY) As TempTileRec
        ReDim Emoticons(0 To MAX_EMOTICONS) As EmoRec
        ReDim MapReport(1 To MAX_MAPS) As MapRec

        ReDim CheckMap(1 To MAX_MAPS) As MapRec
        ReDim MapsAvailable(1 To MAX_MAPS) As Boolean
        
        For i = 1 To MAX_MAPS
            MapsAvailable(i) = False
            ReDim CheckMap(i).Tile(0 To MAX_MAPX, 0 To MAX_MAPY) As TileRec
        Next i

        For i = 0 To MAX_EMOTICONS
            Emoticons(i).Pic = 0
            Emoticons(i).Command = ""
        Next i
        
        Call ClearTempTile
        
        ' Clear out players
        For i = 1 To MAX_PLAYERS
            Call ClearPlayer(i)
        Next i

        frmMirage.Caption = Trim(GAME_NAME)
        App.Title = GAME_NAME
 
        Exit Sub
    End If
        
    ' :::::::::::::::::::
    ' :: Npc hp packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "npchp" Then
        n = Val(Parse(1))
                
        MapNpc(n).HP = Val(Parse(2))
        MapNpc(n).MaxHp = Val(Parse(3))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Alert message packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "alertmsg" Then
        frmMainMenu.Visible = True
        frmMirage.Visible = False
        
        Call TcpDestroy
        frmMainMenu.boxNewChar.Visible = False
        frmMainMenu.boxChars.Visible = False
        frmMainMenu.boxNew.Visible = False
        frmMainMenu.boxLogin.Visible = False
        frmMainMenu.picLogin.Enabled = True
        frmMainMenu.picNewAccount.Enabled = True
        
        frmMainMenu.lblStatus.Caption = Trim(Parse(1))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: All characters packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "allchars" Then
        Call SetStatus("")
        n = 1
        
        frmMainMenu.picLogin.Enabled = False
        frmMainMenu.picNewAccount.Enabled = False
        frmMainMenu.boxChars.Visible = True
        frmMainMenu.lstChars.Clear
        
        For i = 1 To MAX_CHARS
            Name = Parse(n)
            Msg = Parse(n + 1)
            Level = Val(Parse(n + 2))
            
            If Trim(Name) = "" Then
                frmMainMenu.lstChars.AddItem ">Free Slot"
            Else
                frmMainMenu.lstChars.AddItem Name & " a level " & Level & " " & Msg
            End If
            
            n = n + 3
        Next i
        
        frmMainMenu.lstChars.ListIndex = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::::::::::::::
    ' :: Login was successful packet ::
    ' :::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "loginok" Then
        ' Now we can receive game data
        MyIndex = Val(Parse(1))
        
        Call SetStatus("Receiving game data...")
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: New character classes data packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "newcharclasses" Then
        Call SetStatus("")
        
        Max_Classes = Val(Parse(1))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        Class(i).Name = Parse(2)
        Class(i).HP = Val(Parse(3))
        Class(i).MP = Val(Parse(4))
        Class(i).SP = Val(Parse(5))
        Class(i).STR = Val(Parse(6))
        Class(i).DEF = Val(Parse(7))
        Class(i).speed = Val(Parse(8))
        Class(i).MAGI = Val(Parse(9))
        Class(i).MaleSprite = Val(Parse(10))
        Class(i).FemaleSprite = Val(Parse(11))
        Class(i).Locked = Val(Parse(12))
        Class(i).VIT = Val(Parse(13))
        
        frmMainMenu.boxNewChar.Visible = True
        frmMainMenu.txtCharName.SetFocus
        frmMainMenu.txtCharName.SelStart = Len(frmMainMenu.txtCharName.Text)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Classes data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "classesdata" Then
        n = 1
        
        ' Max classes
        Max_Classes = Val(Parse(n))
        ReDim Class(0 To Max_Classes) As ClassRec
        
        n = n + 1
        
        For i = 0 To Max_Classes
            Class(i).Name = Parse(n)
            
            Class(i).HP = Val(Parse(n + 1))
            Class(i).MP = Val(Parse(n + 2))
            Class(i).SP = Val(Parse(n + 3))
            
            Class(i).STR = Val(Parse(n + 4))
            Class(i).DEF = Val(Parse(n + 5))
            Class(i).speed = Val(Parse(n + 6))
            Class(i).MAGI = Val(Parse(n + 7))
            
            Class(i).Locked = Val(Parse(n + 8))
            
            Class(i).VIT = Val(Parse(n + 9))
            
            n = n + 10
        Next i
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: In game packet ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "ingame" Then
        frmMainMenu.Visible = False
        
        InGame = True
        Call GameInit
        Call GameLoop
        If Parse(1) = END_CHAR Then
            MsgBox ("here")
            End
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player inventory packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinv" Then
        n = 1
        For i = 1 To MAX_INV
            Call SetPlayerInvItemNum(MyIndex, i, Val(Parse(n)))
            Call SetPlayerInvItemValue(MyIndex, i, Val(Parse(n + 1)))
            Call SetPlayerInvItemDur(MyIndex, i, Val(Parse(n + 2)))
            
            n = n + 3
        Next i
        Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Player inventory update packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerinvupdate" Then
        n = Val(Parse(1))
        
        Call SetPlayerInvItemNum(MyIndex, n, Val(Parse(2)))
        Call SetPlayerInvItemValue(MyIndex, n, Val(Parse(3)))
        Call SetPlayerInvItemDur(MyIndex, n, Val(Parse(4)))
        Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player worn equipment packet ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerworneq" Then
        Call SetPlayerArmorSlot(MyIndex, Val(Parse(1)))
        Call SetPlayerWeaponSlot(MyIndex, Val(Parse(2)))
        Call SetPlayerHelmetSlot(MyIndex, Val(Parse(3)))
        Call SetPlayerShieldSlot(MyIndex, Val(Parse(4)))
        Call SetPlayerBootsSlot(MyIndex, Val(Parse(5)))
        Call SetPlayerGlovesSlot(MyIndex, Val(Parse(6)))
        Call SetPlayerRingSlot(MyIndex, Val(Parse(7)))
        Call SetPlayerAmuletSlot(MyIndex, Val(Parse(8)))
        Call UpdateVisInv
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Player bank packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerbank" Then
        n = 1
        For i = 1 To MAX_BANK
            Call SetPlayerBankItemNum(MyIndex, i, Val(Parse(n)))
            Call SetPlayerBankItemValue(MyIndex, i, Val(Parse(n + 1)))
            Call SetPlayerBankItemDur(MyIndex, i, Val(Parse(n + 2)))
            
            n = n + 3
        Next i
        
        If frmBank.Visible = True Then Call UpdateBank
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player bank update packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerbankupdate" Then
        n = Val(Parse(1))
        
        Call SetPlayerBankItemNum(MyIndex, n, Val(Parse(2)))
        Call SetPlayerBankItemValue(MyIndex, n, Val(Parse(3)))
        Call SetPlayerBankItemDur(MyIndex, n, Val(Parse(4)))
        If frmBank.Visible = True Then Call UpdateBank
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "openbank" Then
        frmBank.lblBank.Caption = Trim(CheckMap(GetPlayerMap(MyIndex)).Name)
        frmBank.lstInventory.Clear
        frmBank.lstBank.Clear
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(MyIndex, i) > 0 Then
                If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                    frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                        frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name) & " (worn)"
                    Else
                        frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).Name)
                    End If
                End If
            Else
                frmBank.lstInventory.AddItem i & "> Empty"
            End If
            DoEvents
        Next i
        
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(MyIndex, i) > 0 Then
                If Item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
                    frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
                Else
                    If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                        frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).Name) & " (worn)"
                    Else
                        frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).Name)
                    End If
                End If
            Else
                frmBank.lstBank.AddItem i & "> Empty"
            End If
            DoEvents
        Next i
        frmBank.lstBank.ListIndex = 0
        frmBank.lstInventory.ListIndex = 0
        
        frmBank.Show vbModal
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "bankmsg" Then
        frmBank.lblMsg.Caption = Trim(Parse(1))
        Exit Sub
    End If
       
    Dim ShapeW As Long
    ShapeW = 191
        
    ' ::::::::::::::::::::::
    ' :: Player hp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playerhp" Then
        Player(MyIndex).MaxHp = Val(Parse(1))
        Call SetPlayerHP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxHP(MyIndex) > 0 Then
            frmMirage.shpHP.Width = (((GetPlayerHP(MyIndex) / ShapeW) / (GetPlayerMaxHP(MyIndex) / ShapeW)) * ShapeW)
            frmMirage.lblHP.Caption = GetPlayerHP(MyIndex) & " / " & GetPlayerMaxHP(MyIndex)
        End If
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Player mp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playermp" Then
        Player(MyIndex).MaxMP = Val(Parse(1))
        Call SetPlayerMP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxMP(MyIndex) > 0 Then
            frmMirage.shpMP.Width = (((GetPlayerMP(MyIndex) / ShapeW) / (GetPlayerMaxMP(MyIndex) / ShapeW)) * ShapeW)
            frmMirage.lblMP.Caption = GetPlayerMP(MyIndex) & " / " & GetPlayerMaxMP(MyIndex)
        End If
        Exit Sub
    End If
    
    ' speech bubble parse
    If (LCase(Parse(0)) = "mapmsg2") Then
        Bubble(Val(Parse(2))).Text = Parse(1)
        Bubble(Val(Parse(2))).Created = GetTickCount()
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Player sp packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "playersp" Then
        Player(MyIndex).MaxSP = Val(Parse(1))
        Call SetPlayerSP(MyIndex, Val(Parse(2)))
        If GetPlayerMaxSP(MyIndex) > 0 Then
            frmMirage.shpSP.Width = (((GetPlayerSP(MyIndex) / ShapeW) / (GetPlayerMaxSP(MyIndex) / ShapeW)) * ShapeW)
            frmMirage.lblSP.Caption = GetPlayerSP(MyIndex) & " / " & GetPlayerMaxSP(MyIndex)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playerpoints" Then
        Call SetPlayerPOINTS(MyIndex, Val(Parse(1)))
        frmMirage.lblPoints.Caption = Trim(GetPlayerPOINTS(MyIndex))
        frmTraining.Label6.Caption = "Train Skills - " & Val(Parse(1))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Player Stats Packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerstatspacket") Then
        Dim SubStr As Long, SubDef As Long, SubMagi As Long, SubSpeed As Long, SubVit As Long
        SubStr = 0
        SubDef = 0
        SubMagi = 0
        SubSpeed = 0
        SubVit = 0
        
        If GetPlayerWeaponSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).AddSpeed
            SubVit = SubVit + Item(GetPlayerInvItemNum(MyIndex, GetPlayerWeaponSlot(MyIndex))).Data3
        End If
        If GetPlayerArmorSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).AddSpeed
            SubVit = SubVit + Item(GetPlayerInvItemNum(MyIndex, GetPlayerArmorSlot(MyIndex))).Data3
        End If
        If GetPlayerShieldSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).AddSpeed
            SubVit = SubVit + Item(GetPlayerInvItemNum(MyIndex, GetPlayerShieldSlot(MyIndex))).Data3
        End If
        If GetPlayerBootsSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerBootsSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerBootsSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerBootsSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerBootsSlot(MyIndex))).AddSpeed
            SubVit = SubVit + Item(GetPlayerInvItemNum(MyIndex, GetPlayerBootsSlot(MyIndex))).Data3
        End If
        If GetPlayerGlovesSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerGlovesSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerGlovesSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerGlovesSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerGlovesSlot(MyIndex))).AddSpeed
            SubVit = SubVit + Item(GetPlayerInvItemNum(MyIndex, GetPlayerGlovesSlot(MyIndex))).Data3
        End If
        If GetPlayerRingSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).AddSpeed
            SubVit = SubVit + Item(GetPlayerInvItemNum(MyIndex, GetPlayerRingSlot(MyIndex))).Data3
        End If
        If GetPlayerAmuletSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerAmuletSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerAmuletSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerAmuletSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerAmuletSlot(MyIndex))).AddSpeed
            SubVit = SubVit + Item(GetPlayerInvItemNum(MyIndex, GetPlayerAmuletSlot(MyIndex))).Data3
        End If
        If GetPlayerHelmetSlot(MyIndex) > 0 Then
            SubStr = SubStr + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddStr
            SubDef = SubDef + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddDef
            SubMagi = SubMagi + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddMagi
            SubSpeed = SubSpeed + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).AddSpeed
            SubVit = SubVit + Item(GetPlayerInvItemNum(MyIndex, GetPlayerHelmetSlot(MyIndex))).Data3
       End If
        
        If SubStr > 0 Then
            frmTraining.Label1.Caption = "Strength - " & Val(Parse(1)) - SubStr & "/100"
        Else
            frmTraining.Label1.Caption = "Strength - " & Val(Parse(1)) & "/100"
        End If
        If SubDef > 0 Then
            frmTraining.label2.Caption = "Defence - " & Val(Parse(2)) - SubDef & "/100"
        Else
            frmTraining.label2.Caption = "Defence - " & Val(Parse(2)) & "/100"
        End If
        If SubMagi > 0 Then
            frmTraining.Label5.Caption = "Intelligence - " & Val(Parse(4)) - SubMagi & "/100"
        Else
            frmTraining.Label5.Caption = "Intelligence - " & Val(Parse(4)) & "/100"
        End If
        If SubSpeed > 0 Then
            frmTraining.Label4.Caption = "Stamina - " & Val(Parse(3)) - SubSpeed & "/255"
        Else
            frmTraining.Label4.Caption = "Stamina - " & Val(Parse(3)) & "/255"
        End If
        If SubVit > 0 Then
            frmTraining.Label3.Caption = "Vitality - " & Val(Parse(8)) - SubVit & "/255"
        Else
            frmTraining.Label3.Caption = "Vitality - " & Val(Parse(8)) & "/255"
        End If
        
        If SubStr > 0 Then
            frmMirage.lblSTR.Caption = Val(Parse(1)) - SubStr & "  +" & SubStr
        Else
            frmMirage.lblSTR.Caption = Val(Parse(1))
        End If
        If SubDef > 0 Then
            frmMirage.lblDEF.Caption = Val(Parse(2)) - SubDef & "  +" & SubDef
        Else
            frmMirage.lblDEF.Caption = Val(Parse(2))
        End If
        If SubMagi > 0 Then
            frmMirage.lblMAGI.Caption = Val(Parse(4)) - SubMagi & "  +" & SubMagi
        Else
            frmMirage.lblMAGI.Caption = Val(Parse(4))
        End If
        If SubSpeed > 0 Then
            frmMirage.lblSPEED.Caption = Val(Parse(3)) - SubSpeed & "  +" & SubSpeed
        Else
            frmMirage.lblSPEED.Caption = Val(Parse(3))
        End If
        Call SetPlayerVIT(MyIndex, Val(Parse(8)))
        If SubVit > 0 Then
            frmMirage.lblVit.Caption = Val(Parse(8)) - SubVit & "  +" & SubVit
        Else
            frmMirage.lblVit.Caption = Val(Parse(8))
        End If
        
        frmMirage.lblEXP.Caption = Val(Parse(6)) & " / " & Val(Parse(5))
        frmMirage.shpTNL.Width = (((Val(Parse(6)) / ShapeW) / (Val(Parse(5)) / ShapeW)) * ShapeW)
        frmMirage.lblLevel.Caption = "" & Val(Parse(7))
        Player(MyIndex).Level = Val(Parse(7))
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playerclass" Then
        Call SetPlayerClass(MyIndex, Val(Parse(1)))
        Exit Sub
    End If
                
    ' ::::::::::::::::::::::::
    ' :: Player data packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdata" Then
        i = Val(Parse(1))
        Call SetPlayerName(i, Parse(2))
        Call SetPlayerSprite(i, Val(Parse(3)))
        Call SetPlayerMap(i, Val(Parse(4)))
        Call SetPlayerX(i, Val(Parse(5)))
        Call SetPlayerY(i, Val(Parse(6)))
        Call SetPlayerDir(i, Val(Parse(7)))
        Call SetPlayerAccess(i, Val(Parse(8)))
        Call SetPlayerPK(i, Val(Parse(9)))
        Call SetPlayerGuild(i, Parse(10))
        Call SetPlayerGuildAccess(i, Val(Parse(11)))
        Call SetPlayerClass(i, Val(Parse(12)))
        
        ' Make sure they aren't walking
        Player(i).Moving = 0
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        
        ' Check if the player is the client player, and if so reset directions
        If i = MyIndex Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
        
        If i = MyIndex Then
            frmMirage.lblName.Caption = Trim(GetPlayerName(MyIndex))
            frmMirage.lblClass.Caption = Trim(Class(GetPlayerClass(MyIndex)).Name)
            If Val(Parse(13)) = 0 Then
                Player(MyIndex).Sex = SEX_MALE
                frmMirage.lblGender.Caption = "Male"
            Else
                Player(MyIndex).Sex = SEX_FEMALE
                frmMirage.lblGender.Caption = "Female"
            End If
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Player movement packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playermove") Then
        i = Val(Parse(1))
        x = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        Call SetPlayerX(i, x)
        Call SetPlayerY(i, Y)
        Call SetPlayerDir(i, Dir)
                
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = n
        
        Select Case GetPlayerDir(i)
            Case DIR_UP
                Player(i).YOffset = PIC_Y
            Case DIR_DOWN
                Player(i).YOffset = PIC_Y * -1
            Case DIR_LEFT
                Player(i).XOffset = PIC_X
            Case DIR_RIGHT
                Player(i).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Npc movement packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcmove") Then
        i = Val(Parse(1))
        x = Val(Parse(2))
        Y = Val(Parse(3))
        Dir = Val(Parse(4))
        n = Val(Parse(5))

        MapNpc(i).x = x
        MapNpc(i).Y = Y
        MapNpc(i).Dir = Dir
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = n
        
        Select Case MapNpc(i).Dir
            Case DIR_UP
                MapNpc(i).YOffset = PIC_Y
            Case DIR_DOWN
                MapNpc(i).YOffset = PIC_Y * -1
            Case DIR_LEFT
                MapNpc(i).XOffset = PIC_X
            Case DIR_RIGHT
                MapNpc(i).XOffset = PIC_X * -1
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Player direction packet ::
    ' :::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        Call SetPlayerDir(i, Dir)
        
        Player(i).XOffset = 0
        Player(i).YOffset = 0
        Player(i).Moving = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: NPC direction packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcdir") Then
        i = Val(Parse(1))
        Dir = Val(Parse(2))
        MapNpc(i).Dir = Dir
        
        MapNpc(i).XOffset = 0
        MapNpc(i).YOffset = 0
        MapNpc(i).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Player XY location packet ::
    ' :::::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "playerxy") Then
        x = Val(Parse(1))
        Y = Val(Parse(2))
        
        Call SetPlayerX(MyIndex, x)
        Call SetPlayerY(MyIndex, Y)
        
        ' Make sure they aren't walking
        Player(MyIndex).Moving = 0
        Player(MyIndex).XOffset = 0
        Player(MyIndex).YOffset = 0
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "attack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        Player(i).Attacking = 1
        Player(i).AttackTimer = GetTickCount
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: NPC attack packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npcattack") Then
        i = Val(Parse(1))
        
        ' Set player to attacking
        MapNpc(i).Attacking = 1
        MapNpc(i).AttackTimer = GetTickCount
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Check for map packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "checkformap") Then
        ' Erase all players except self
        For i = 1 To MAX_PLAYERS
            If i <> MyIndex Then
                Call SetPlayerMap(i, 0)
            End If
        Next i
        
        ' Erase all temporary tile values
        Call ClearTempTile

        ' Get map num
        x = Val(Parse(1))
        
        ' Get revision
        Y = Val(Parse(2))
        
        If FileExist("maps\map" & x & ".dat") Then
            If GetMapRevision(x) = Y Then
                Call LoadMap(x)
                
                Call SendData("needmap" & SEP_CHAR & "no" & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
                
        ' Either the revisions didn't match or we dont have the map, so we need it
        'Call SendMapData("getmap" & SEP_CHAR & X & SEP_CHAR & END_CHAR)
        Call SendData("needmap" & SEP_CHAR & "yes" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Map send completed packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdone" Then
        CheckMap(GetPlayerMap(MyIndex)) = SaveMap
        MapsAvailable(GetPlayerMap(MyIndex)) = True
        
        For i = 1 To MAX_MAP_ITEMS
            MapItem(i) = SaveMapItem(i)
        Next i
        
        For i = 1 To MAX_MAP_NPCS
            MapNpc(i) = SaveMapNpc(i)
        Next i
        
        GettingMap = False
        
        ' Play music
        If Trim(CheckMap(GetPlayerMap(MyIndex)).Music) <> "None" Then
            If ReadINI("CONFIG", "Music", App.Path & "\config.ini") = 1 Then
                Call PlayMidi(Trim(CheckMap(GetPlayerMap(MyIndex)).Music))
            Else
                Call StopMidi
            End If
        Else
            Call StopMidi
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "mapdata" Then
        n = 1
        
        SaveMap.Name = Parse(n + 1)
        SaveMap.Revision = Val(Parse(n + 2))
        SaveMap.Moral = Val(Parse(n + 3))
        SaveMap.Up = Val(Parse(n + 4))
        SaveMap.Down = Val(Parse(n + 5))
        SaveMap.Left = Val(Parse(n + 6))
        SaveMap.Right = Val(Parse(n + 7))
        SaveMap.Music = Parse(n + 8)
        SaveMap.BootMap = Val(Parse(n + 9))
        SaveMap.BootX = Val(Parse(n + 10))
        SaveMap.BootY = Val(Parse(n + 11))
        SaveMap.Indoor = Val(Parse(n + 12))
        SaveMap.Random = 0
        
        n = n + 13
        
        For Y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                SaveMap.Tile(x, Y).Ground = Val(Parse(n))
                SaveMap.Tile(x, Y).Mask = Val(Parse(n + 1))
                SaveMap.Tile(x, Y).Anim = Val(Parse(n + 2))
                SaveMap.Tile(x, Y).Mask2 = Val(Parse(n + 3))
                SaveMap.Tile(x, Y).M2Anim = Val(Parse(n + 4))
                SaveMap.Tile(x, Y).Fringe = Val(Parse(n + 5))
                SaveMap.Tile(x, Y).FAnim = Val(Parse(n + 6))
                SaveMap.Tile(x, Y).Fringe2 = Val(Parse(n + 7))
                SaveMap.Tile(x, Y).F2Anim = Val(Parse(n + 8))
                SaveMap.Tile(x, Y).Type = Val(Parse(n + 9))
                SaveMap.Tile(x, Y).Data1 = Val(Parse(n + 10))
                SaveMap.Tile(x, Y).Data2 = Val(Parse(n + 11))
                SaveMap.Tile(x, Y).Data3 = Val(Parse(n + 12))
                SaveMap.Tile(x, Y).String1 = Parse(n + 13)
                SaveMap.Tile(x, Y).String2 = Parse(n + 14)
                SaveMap.Tile(x, Y).String3 = Parse(n + 15)
                
                n = n + 16
            Next x
        Next Y
        
        For x = 1 To MAX_MAP_NPCS
            SaveMap.Npc(x) = Val(Parse(n))
            n = n + 1
        Next x
                
        ' Save the map
        Call SaveLocalMap(Val(Parse(1)))
        
        ' Check if we get a map from someone else and if we were editing a map cancel it out
        If InEditor Then
            InEditor = False
            frmMapEditor.Visible = False
            'frmmirage.show
            'frmMirage.picMapEditor.Visible = False
            
            If frmMapWarp.Visible Then
                Unload frmMapWarp
            End If
            
            If frmMapProperties.Visible Then
                Unload frmMapProperties
            End If
        End If
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::::
    ' :: Map items data packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapitemdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_ITEMS
            SaveMapItem(i).Num = Val(Parse(n))
            SaveMapItem(i).Value = Val(Parse(n + 1))
            SaveMapItem(i).Dur = Val(Parse(n + 2))
            SaveMapItem(i).x = Val(Parse(n + 3))
            SaveMapItem(i).Y = Val(Parse(n + 4))
            
            n = n + 5
        Next i
        
        Exit Sub
    End If
        
    ' :::::::::::::::::::::::::
    ' :: Map npc data packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapnpcdata" Then
        n = 1
        
        For i = 1 To MAX_MAP_NPCS
            SaveMapNpc(i).Num = Val(Parse(n))
            SaveMapNpc(i).x = Val(Parse(n + 1))
            SaveMapNpc(i).Y = Val(Parse(n + 2))
            SaveMapNpc(i).Dir = Val(Parse(n + 3))
            
            n = n + 4
        Next i
        
        Exit Sub
    End If

    If LCase(Parse(0)) = "broadcastmsg" Then
        If frmMirage.chkBroadcast.Value = Checked Then
            Call AddText(Parse(1), Val(Parse(2)))
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "saymsg") Or (LCase(Parse(0)) = "globalmsg") Or (LCase(Parse(0)) = "playermsg") Or (LCase(Parse(0)) = "mapmsg") Or (LCase(Parse(0)) = "adminmsg") Then
        Call AddText(Parse(1), Val(Parse(2)))
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Item spawn packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnitem" Then
        n = Val(Parse(1))
        
        MapItem(n).Num = Val(Parse(2))
        MapItem(n).Value = Val(Parse(3))
        MapItem(n).Dur = Val(Parse(4))
        MapItem(n).x = Val(Parse(5))
        MapItem(n).Y = Val(Parse(6))
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Item editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "itemeditor") Then
        InItemsEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_ITEMS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Item(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update item packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateitem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).StrReq = Val(Parse(8))
        Item(n).DefReq = Val(Parse(9))
        Item(n).SpeedReq = Val(Parse(10))
        Item(n).ClassReq = Val(Parse(11))
        Item(n).AccessReq = Val(Parse(12))
        
        Item(n).AddHP = Val(Parse(13))
        Item(n).AddMP = Val(Parse(14))
        Item(n).AddSP = Val(Parse(15))
        Item(n).AddStr = Val(Parse(16))
        Item(n).AddDef = Val(Parse(17))
        Item(n).AddMagi = Val(Parse(18))
        Item(n).AddSpeed = Val(Parse(19))
        Item(n).AddEXP = Val(Parse(20))
        
        Item(n).desc = Parse(21)
        Exit Sub
    End If
       
    ' ::::::::::::::::::::::
    ' :: Edit item packet :: <- Used for item editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "edititem") Then
        n = Val(Parse(1))
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        Item(n).StrReq = Val(Parse(8))
        Item(n).DefReq = Val(Parse(9))
        Item(n).SpeedReq = Val(Parse(10))
        Item(n).ClassReq = Val(Parse(11))
        Item(n).AccessReq = Val(Parse(12))
        
        Item(n).AddHP = Val(Parse(13))
        Item(n).AddMP = Val(Parse(14))
        Item(n).AddSP = Val(Parse(15))
        Item(n).AddStr = Val(Parse(16))
        Item(n).AddDef = Val(Parse(17))
        Item(n).AddMagi = Val(Parse(18))
        Item(n).AddSpeed = Val(Parse(19))
        Item(n).AddEXP = Val(Parse(20))
        
        Item(n).desc = Parse(21)
        
        ' Initialize the item editor
        Call ItemEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Npc spawn packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "spawnnpc" Then
        n = Val(Parse(1))
        
        MapNpc(n).Num = Val(Parse(2))
        MapNpc(n).x = Val(Parse(3))
        MapNpc(n).Y = Val(Parse(4))
        MapNpc(n).Dir = Val(Parse(5))
        MapNpc(n).Big = Val(Parse(6))
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Npc dead packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "npcdead" Then
        n = Val(Parse(1))
        
        MapNpc(n).Num = 0
        MapNpc(n).x = 0
        MapNpc(n).Y = 0
        MapNpc(n).Dir = 0
        
        ' Client use only
        MapNpc(n).XOffset = 0
        MapNpc(n).YOffset = 0
        MapNpc(n).Moving = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Npc editor packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "npceditor") Then
        InNpcEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_NPCS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Npc(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Update npc packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatenpc") Then
        n = Val(Parse(1))
        
        ' Update the item
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = ""
        Npc(n).Sprite = Val(Parse(3))
        Npc(n).SpawnSecs = 0
        Npc(n).Behavior = 0
        Npc(n).Range = 0
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = 0
            Npc(n).ItemNPC(i).ItemNum = 0
            Npc(n).ItemNPC(i).ItemValue = 0
        Next i
        Npc(n).STR = 0
        Npc(n).DEF = 0
        Npc(n).speed = 0
        Npc(n).MAGI = 0
        Npc(n).Big = Val(Parse(4))
        Npc(n).MaxHp = Val(Parse(5))
        Npc(n).EXP = 0
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit npc packet :: <- Used for item editor admins only
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editnpc") Then
        n = Val(Parse(1))
        
        ' Update the npc
        Npc(n).Name = Parse(2)
        Npc(n).AttackSay = Parse(3)
        Npc(n).Sprite = Val(Parse(4))
        Npc(n).SpawnSecs = Val(Parse(5))
        Npc(n).Behavior = Val(Parse(6))
        Npc(n).Range = Val(Parse(7))
        Npc(n).STR = Val(Parse(8))
        Npc(n).DEF = Val(Parse(9))
        Npc(n).speed = Val(Parse(10))
        Npc(n).MAGI = Val(Parse(11))
        Npc(n).Big = Val(Parse(12))
        Npc(n).MaxHp = Val(Parse(13))
        Npc(n).EXP = Val(Parse(14))
        
        z = 15
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = Val(Parse(z))
            Npc(n).ItemNPC(i).ItemNum = Val(Parse(z + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(Parse(z + 2))
            z = z + 3
        Next i
        
        ' Initialize the npc editor
        Call NpcEditorInit

        Exit Sub
    End If

    ' ::::::::::::::::::::
    ' :: Map key packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "mapkey") Then
        x = Val(Parse(1))
        Y = Val(Parse(2))
        n = Val(Parse(3))
                
        TempTile(x, Y).DoorOpen = n
        
        Exit Sub
    End If

    ' :::::::::::::::::::::
    ' :: Edit map packet ::
    ' :::::::::::::::::::::
    If (LCase(Parse(0)) = "editmap") Then
        Call EditorInit
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Shop editor packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "shopeditor") Then
        InShopEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SHOPS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Shop(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update shop packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updateshop") Then
        n = Val(Parse(1))
        
        ' Update the shop name
        Shop(n).Name = Parse(2)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet :: <- Used for shop editor admins only
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "editshop") Then
        ShopNum = Val(Parse(1))
        
        ' Update the shop
        Shop(ShopNum).Name = Parse(2)
        Shop(ShopNum).JoinSay = Parse(3)
        Shop(ShopNum).LeaveSay = Parse(4)
        Shop(ShopNum).FixesItems = Val(Parse(5))
        
        n = 6
        For i = 1 To MAX_TRADES
            
            GiveItem = Val(Parse(n))
            GiveValue = Val(Parse(n + 1))
            GetItem = Val(Parse(n + 2))
            GetValue = Val(Parse(n + 3))
            
            Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
            Shop(ShopNum).TradeItem(i).GiveValue = GiveValue
            Shop(ShopNum).TradeItem(i).GetItem = GetItem
            Shop(ShopNum).TradeItem(i).GetValue = GetValue
            
            n = n + 4
        Next i
        
        ' Initialize the shop editor
        Call ShopEditorInit

        Exit Sub
    End If

    ' :::::::::::::::::::::::::
    ' :: Spell editor packet ::
    ' :::::::::::::::::::::::::
    If (LCase(Parse(0)) = "spelleditor") Then
        InSpellEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        ' Add the names
        For i = 1 To MAX_SPELLS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Spell(i).Name)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Update spell packet ::
    ' ::::::::::::::::::::::::
    If (LCase(Parse(0)) = "updatespell") Then
        n = Val(Parse(1))
        
        ' Update the spell name
        Spell(n).Name = Parse(2)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet :: <- Used for spell editor admins only
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "editspell") Then
        n = Val(Parse(1))
        
        ' Update the spell
        Spell(n).Name = Parse(2)
        Spell(n).ClassReq = Val(Parse(3))
        Spell(n).LevelReq = Val(Parse(4))
        Spell(n).Type = Val(Parse(5))
        Spell(n).Data1 = Val(Parse(6))
        Spell(n).Data2 = Val(Parse(7))
        Spell(n).Data3 = Val(Parse(8))
        Spell(n).MPCost = Val(Parse(9))
        Spell(n).Sound = Val(Parse(10))
                        
        ' Initialize the spell editor
        Call SpellEditorInit

        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Trade packet ::
    ' ::::::::::::::::::
    If (LCase(Parse(0)) = "trade") Then
        ShopNum = Val(Parse(1))
        
        frmTrade.lblShopName.Caption = Trim(Shop(ShopNum).Name)
        
        n = 3
        For i = 1 To MAX_TRADES
            GiveItem = Val(Parse(n))
            GiveValue = Val(Parse(n + 1))
            GetItem = Val(Parse(n + 2))
            GetValue = Val(Parse(n + 3))
            ItemGetS(i) = GetItem
            ItemGiveS(i) = GiveItem
            ItemGetSS(i) = GetValue
            ItemGiveSS(i) = GiveValue
            
            If ItemGetS(i) > 0 Then
                frmTrade.Label0(i - 1).Visible = True
                frmTrade.ItemS(i - 1).Visible = True
                frmTrade.Price(i - 1).Visible = True
                frmTrade.Deal(i - 1).Visible = True
                frmTrade.ItemS(i - 1).Caption = ItemGetSS(i) & " " & Trim(Item(ItemGetS(i)).Name)
                frmTrade.Price(i - 1).Caption = ItemGiveSS(i) & " " & Trim(Item(ItemGiveS(i)).Name)
            Else
                frmTrade.Label0(i - 1).Visible = False
                frmTrade.ItemS(i - 1).Visible = False
                frmTrade.Price(i - 1).Visible = False
                frmTrade.Deal(i - 1).Visible = False
                frmTrade.ItemS(i - 1).Caption = ""
                frmTrade.Price(i - 1).Caption = ""
            End If
            
            n = n + 4
        Next i
        
        frmTrade.Show vbModeless, frmMirage
        Exit Sub
    End If

    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If (LCase(Parse(0)) = "spells") Then
        
        frmMirage.picPlayerSpells.ZOrder (0)
        frmMirage.lstSpells.Clear
        
        ' Put spells known in player record
        For i = 1 To MAX_PLAYER_SPELLS
            Player(MyIndex).Spell(i) = Val(Parse(i))
            If Player(MyIndex).Spell(i) <> 0 Then
                frmMirage.lstSpells.AddItem i & ": " & Trim(Spell(Player(MyIndex).Spell(i)).Name)
            Else
                frmMirage.lstSpells.AddItem "<free spells slot>"
            End If
        Next i
        
        frmMirage.lstSpells.ListIndex = 0
    End If

    ' ::::::::::::::::::::
    ' :: Weather packet ::
    ' ::::::::::::::::::::
    If (LCase(Parse(0)) = "weather") Then
        If Val(Parse(1)) = WEATHER_RAINING And GameWeather <> WEATHER_RAINING Then
            Call AddText("You see drops of rain falling from the sky above!", BrightGreen)
        End If
        If Val(Parse(1)) = WEATHER_THUNDER And GameWeather <> WEATHER_THUNDER Then
            Call AddText("You see lightning in the sky above!", BrightGreen)
        End If
        If Val(Parse(1)) = WEATHER_SNOWING And GameWeather <> WEATHER_SNOWING Then
            Call AddText("You see snow falling from the sky above!", BrightGreen)
        End If
        
        If Val(Parse(1)) = WEATHER_NONE Then
            If GameWeather = WEATHER_RAINING Then
                Call AddText("The rain beings to calm.", BrightGreen)
            ElseIf GameWeather = WEATHER_SNOWING Then
                Call AddText("The snow is melting away.", BrightGreen)
            ElseIf GameWeather = WEATHER_THUNDER Then
                Call AddText("The lightning begins to disapear.", BrightGreen)
            End If
        End If
        GameWeather = Val(Parse(1))
        RainIntensity = Val(Parse(2))
        If MAX_RAINDROPS <> RainIntensity Then
            MAX_RAINDROPS = RainIntensity
            ReDim DropRain(1 To MAX_RAINDROPS) As DropRainRec
            ReDim DropSnow(1 To MAX_RAINDROPS) As DropRainRec
        End If
    End If

    ' :::::::::::::::::
    ' :: Time packet ::
    ' :::::::::::::::::
    If (LCase(Parse(0)) = "time") Then
        GameTime = Val(Parse(1))
        If GameTime = TIME_NIGHT Then
            Call AddText("The night's shadow has fallen upon us!", White)
        Else
            Call AddText("Day has dawned!", White)
        End If
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Get Online List ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "onlinelist" Then
    frmMirage.lstOnline.Clear
    
        n = 2
        z = Val(Parse(1))
        For x = n To (z + 1)
            frmMirage.lstOnline.AddItem Trim(Parse(n))
            n = n + 2
        Next x
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::
    ' :: Retrieve the player's inventory ::
    ' :::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "pptrading" Then
        frmPlayerTrade.Items1.Clear
        frmPlayerTrade.Items2.Clear
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = ""
            Trading2(i).InvNum = 0
            Trading2(i).InvName = ""
            frmPlayerTrade.Items1.AddItem i & ": <Nothing>"
            frmPlayerTrade.Items2.AddItem i & ": <Nothing>"
        Next i
        
        frmPlayerTrade.Items1.ListIndex = 0
        
        Call UpdateTradeInventory
        frmPlayerTrade.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "qtrade" Then
        For i = 1 To MAX_PLAYER_TRADES
            Trading(i).InvNum = 0
            Trading(i).InvName = ""
            Trading2(i).InvNum = 0
            Trading2(i).InvName = ""
        Next i
        
        frmPlayerTrade.Command1.ForeColor = &H0&
        frmPlayerTrade.Command2.ForeColor = &H0&
        
        frmPlayerTrade.Visible = False
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "updatetradeitem" Then
            n = Val(Parse(1))
            
            Trading2(n).InvNum = Val(Parse(2))
            Trading2(n).InvName = Parse(3)
            
            If STR(Trading2(n).InvNum) <= 0 Then
                frmPlayerTrade.Items2.List(n - 1) = n & ": <Nothing>"
            Else
                frmPlayerTrade.Items2.List(n - 1) = n & ": " & Trim(Trading2(n).InvName)
            End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Stop trading with player ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "trading" Then
        n = Val(Parse(1))
            If n = 0 Then frmPlayerTrade.Command2.ForeColor = &H0&
            If n = 1 Then frmPlayerTrade.Command2.ForeColor = &HFF00&
        Exit Sub
    End If
    
' :::::::::::::::::::::::::
' :: Chat System Packets ::
' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "ppchatting" Then
        frmPlayerChat.txtChat.Text = ""
        frmPlayerChat.txtSay.Text = ""
        Player(Val(Parse(1))).Name = Trim(Parse(2))
        frmPlayerChat.Label1.Caption = "Chatting With: " & Trim(Player(Val(Parse(1))).Name)

        frmPlayerChat.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "qchat" Then
        frmPlayerChat.txtChat.Text = ""
        frmPlayerChat.txtSay.Text = ""
        frmPlayerChat.Visible = False
        frmPlayerTrade.Command2.ForeColor = &H8000000F
        frmPlayerTrade.Command1.ForeColor = &H8000000F
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "sendchat" Then
        Dim s As String
  
        s = vbNewLine & GetPlayerName(Val(Parse(2))) & "> " & Trim(Parse(1))
        frmPlayerChat.txtChat.SelStart = Len(frmPlayerChat.txtChat.Text)
        frmPlayerChat.txtChat.SelColor = QBColor(Brown)
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
    If LCase(Parse(0)) = "sound" Then
        s = LCase(Parse(1))
        Select Case Trim(s)
            Case "attack"
                Call PlaySound("sword.wav")
            Case "critical"
                Call PlaySound("critical.wav")
            Case "miss"
                Call PlaySound("miss.wav")
            Case "key"
                Call PlaySound("key.wav")
            Case "magic"
                Call PlaySound("magic" & Val(Parse(2)) & ".wav")
            Case "warp"
                Call PlaySound("warp.wav")
            Case "pain"
                Call PlaySound("pain.wav")
            Case "soundattribute"
                Call PlaySound(Trim(Parse(2)))
        End Select
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::
    ' :: Sprite Change Confirmation Packet ::
    ' :::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "spritechange" Then
        i = MsgBox("Would you like to buy this sprite?", 4, "Buying Sprite")
        If i = 6 Then
            Call SendData("buysprite" & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::
    ' :: Change Player Direction Packet ::
    ' ::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "changedir" Then
        Player(Val(Parse(2))).Dir = Val(Parse(1))
        Exit Sub
    End If
        
    ' :::::::::::::::::::
    ' :: Prompt Packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "prompt" Then
        i = MsgBox(Trim(Parse(1)), vbYesNo)
        Call SendData("prompt" & SEP_CHAR & i & SEP_CHAR & Val(Parse(2)) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

    ' ::::::::::::::::::::::::::::
    ' :: Emoticon editor packet ::
    ' ::::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "emoticoneditor") Then
        InEmoticonEditor = True
        
        frmIndex.Show
        frmIndex.lstIndex.Clear
        
        For i = 0 To MAX_EMOTICONS
            frmIndex.lstIndex.AddItem i & ": " & Trim(Emoticons(i).Command)
        Next i
        
        frmIndex.lstIndex.ListIndex = 0
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "updateemoticon") Then
        n = Val(Parse(1))
        
        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        Exit Sub
    End If

    If (LCase(Parse(0)) = "editemoticon") Then
        n = Val(Parse(1))

        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        
        Call EmoticonEditorInit
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "updateemoticon") Then
        n = Val(Parse(1))
        
        Emoticons(n).Command = Parse(2)
        Emoticons(n).Pic = Val(Parse(3))
        Exit Sub
    End If

    If (LCase(Parse(0)) = "checkemoticons") Then
        n = Val(Parse(1))
        
        Player(n).Emoticon = Val(Parse(2))
        Player(n).EmoticonT = GetTickCount
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "checksprite") Then
        n = Val(Parse(1))
        
        Player(n).Sprite = Val(Parse(2))
        Exit Sub
    End If
    
    If (LCase(Parse(0)) = "mapreport") Then
        n = 1
        
        frmMapReport.lstIndex.Clear
        For i = 1 To MAX_MAPS
            frmMapReport.lstIndex.AddItem i & ": " & Trim(Parse(n))
            n = n + 1
        Next i
        
        frmMapReport.Show vbModeless, frmMirage
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "damagedisplay" Then
        For i = 1 To MAX_BLT_LINE
            If BattleMsg(i).Index <= 0 Then
                BattleMsg(i).Index = Val(Parse(1))
                BattleMsg(i).Msg = Parse(2)
                BattleMsg(i).Color = Val(Parse(3))
                BattleMsg(i).Time = GetTickCount
                BattleMsg(i).Done = 1
                BattleMsg(i).Y = 0
                Exit Sub
            Else
                BattleMsg(i).Y = BattleMsg(i).Y - 15
            End If
        Next i
        
        z = 1
        For i = 1 To MAX_BLT_LINE
            If i < MAX_BLT_LINE Then
                If BattleMsg(i).Y < BattleMsg(i + 1).Y Then z = i
            Else
                If BattleMsg(i).Y < BattleMsg(1).Y Then z = i
            End If
        Next i
                    
        BattleMsg(z).Index = Val(Parse(1))
        BattleMsg(z).Msg = Parse(2)
        BattleMsg(z).Color = Val(Parse(3))
        BattleMsg(z).Time = GetTickCount
        BattleMsg(z).Done = 1
        BattleMsg(z).Y = 0
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "partydisplay" Then
        For i = 0 To MAX_PARTY_MEMS
            PartyMems(i).Name = ""
            PartyMems(i).Y = 0
        Next i
        For i = 0 To MAX_PARTY_MEMS
            If Parse(1 + i) = "" Then
                Exit Sub
            Else
                PartyMems(i).Name = Parse(1 + i)
                PartyMems(i).Y = (i + 1) * 15
            End If
        Next i
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "noparty" Then
        For i = 0 To MAX_PARTY_MEMS
            PartyMems(i).Name = ""
            PartyMems(i).Y = 0
        Next i
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "traininghouse" Then
        frmTraining.Show vbModal
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "reloadmaps" Then
        'Call SendMapData("reloadmap" & SEP_CHAR & Val(Parse(1)) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
End Sub
