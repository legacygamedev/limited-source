Attribute VB_Name = "modHandleData"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' -- Parses and handles String packets    --
' ------------------------------------------

Public Sub HandleData(ByVal Data As String)
Dim Parse() As String

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    
    ' Add the data to the debug window if we are in debug mode
    If DEBUG_MODE Then
        If Not frmDebug.Visible Then frmDebug.Visible = True
        TextAdd frmDebug.txtDebug, "((( Processed Packet " & Parse(0) & " )))"
    End If
    
    Select Case Parse(0)
    
        Case SAlertMsg
            HandleAlertMsg Parse
        Case SAllChars
            HandleAllChars Parse
        Case SLoginOk
            HandleLoginOk Parse
        Case SNewCharClasses
            HandleNewCharClasses Parse
        Case SClassesData
            HandleClassesData Parse
        Case SInGame
            HandleInGame
        Case SPlayerInv
            HandlePlayerInv Parse
        Case SPlayerInvUpdate
            HandlePlayerInvUpdate Parse
        Case SPlayerWornEq
            HandlePlayerWornEq Parse
        Case SPlayerHp
            HandlePlayerHp Parse
        Case SPlayerMp
            HandlePlayerMp Parse
        Case SPlayerSp
            HandlePlayerSp Parse
        Case SPlayerStats
            HandlePlayerStats Parse
        Case SPlayerData
            HandlePlayerData Parse
        Case SPlayerMove
            HandlePlayerMove Parse
        Case SNpcMove
            HandleNpcMove Parse
        Case SPlayerDir
            HandlePlayerDir Parse
        Case SNpcDir
            HandleNpcDir Parse
        Case SPlayerXY
            HandlePlayerXY Parse
        Case SAttack
            HandleAttack Parse
        Case SNpcAttack
            HandleNpcAttack Parse
        Case SCheckForMap
            HandleCheckForMap Parse
        Case SMapData
            HandleMapData Parse
        Case SMapItemData
            HandleMapItemData Parse
        Case SMapNpcData
            HandleMapNpcData Parse
        Case SMapDone
            HandleMapDone
        Case SMessage
            HandleMessage Parse
        Case SGlobalMsg
            HandleGlobalMsg Parse
        Case SAdminMsg
            HandleAdminMsg Parse
        Case SPlayerMsg
            HandlePlayerMsg Parse
        Case SMapMsg
            HandleMapMsg Parse
        Case SSpawnItem
            HandleSpawnItem Parse
        Case SItemEditor
            HandleItemEditor
        Case SUpdateItem
            HandleUpdateItem Parse
        Case SEditItem
            HandleEditItem Parse
        Case SSpawnNpc
            HandleSpawnNpc Parse
        Case SNpcDead
            HandleNpcDead Parse
        Case SNpcEditor
            HandleNpcEditor
        Case SUpdateNpc
            HandleUpdateNpc Parse
        Case SEditNpc
            HandleEditNpc Parse
        Case SMapKey
            HandleMapKey Parse
        Case SEditMap
            HandleEditMap
        Case SShopEditor
            HandleShopEditor
        Case SUpdateShop
            HandleUpdateShop Parse
        Case SEditShop
            HandleEditShop Parse
        Case SREditor
            HandleRefresh
        Case SSpellEditor
            HandleSpellEditor
        Case SUpdateSpell
            HandleUpdateSpell Parse
        Case SEditSpell
            HandleEditSpell Parse
        Case STrade
            HandleTrade Parse
        Case SSpells
            HandleSpells Parse
        Case SLeft
            HandleLeft Parse
        Case SConfigPass
            HandleConfigPass Parse
        Case SGameOptions
            HandleGameOptions Parse
        Case SAnimation
            HandleAnimation Parse
        Case SSoundPlay
            HandleSoundPlay Parse
        Case SPlayerPoints
            HandlePlayerPoints Parse
        Case SPlayerLevel
            HandlePlayerLevel Parse
        Case SClassName
            HandleClassName Parse
        Case SPlayerStatBuffs
            HandlePlayerStatBuffs Parse
        Case SSignEditor
            HandleSignEditor
        Case SEditSign
            HandleEditSign Parse
        Case SUpdateSign
            HandleUpdateSign Parse
        Case SScrollingText
            HandleScrollingText Parse
        Case SGuildCreation
            HandleGuildCreation
        Case SPlayerGuild
            HandlePlayerGuild Parse
        Case SGuildInvite
            HandleGuildInvite Parse
        Case SAnimEditor
            HandleAnimEditor
        Case SEditAnim
            HandleEditAnim Parse
        Case SUpdateAnim
            HandleUpdateAnim Parse
        Case SPing
            HandlePing
        Case SNpcHP
            HandleNpcHP Parse
        Case SNormalMsg
            HandleNormalMsg Parse
        Case SCastSuccess
            HandleCastSuccess Parse
        Case SExpUpdate
            HandleExpUpdate Parse
    End Select
    
End Sub

' ::::::::::::::::::::::::::
 ' :: Alert message packet ::
 ' ::::::::::::::::::::::::::
Private Sub HandleAlertMsg(ByRef Parse() As String)

    CurrentWindow = Window_State.Main_Menu
    
    frmStatus.Visible = False
    frmMainMenu.Visible = True
    
    MsgBox Parse(1), vbOKOnly, Game_Name
    
End Sub

' :::::::::::::::::::::::::::
 ' :: All characters packet ::
 ' :::::::::::::::::::::::::::
Private Sub HandleAllChars(ByRef Parse() As String)
Dim n As Long
Dim i As Long
Dim Level As Long
Dim Name As String
Dim Msg As String

    n = 1
    
    'SetStatus "Initializing DirectX 7..."
    
    ' load up DX7 before they get to see their characters
    'InitDirectDraw
    'InitSurfaces
    
    ' getting the width of all tile sheets
    'With frmMainGame.picCheckSize
    '   For i = 0 To MAX_TILESETS
    '      .Picture = LoadPicture(App.Path & GFX_PATH & "tiles" & i & GFX_EXT)
    '      LoadTileSheetWidth i, .Width
    '      .Picture = LoadPicture()
    '      .Width = 1
    '      .Height = 1
    '   Next
    'End With
    
    'If Sound_On Then InitDirectSound
    'If Music_On Then InitDirectMusic
    
    ' late binding
    'If DX7 Is Nothing Then Set DX7 = New DirectX7
    
    CurrentWindow = Window_State.Chars
    
    frmChars.Caption = Game_Name & " [Select Character]"
    frmChars.Visible = True
    frmStatus.Visible = False
    
    Char_Selected = 1
    
    For i = 1 To MAX_CHARS
        
        If i = 1 Then
            frmChars.picCharBox(i).Picture = LoadPicture(App.Path & GFX_PATH & "Interface/char1selected.bmp")
        Else
            frmChars.picCharBox(i).Picture = LoadPicture(App.Path & GFX_PATH & "Interface/char" & i & "unselected.bmp")
        End If
        
        Name = Parse(n)
        Msg = Parse(n + 1)
        Level = CLng(Parse(n + 2))
        
        If LenB(Trim$(Name)) > 0 Then
            frmChars.lblCharDetails(i).Caption = "[" & Name & "]" & vbNewLine & _
                                                 "Level " & Level & " " & Msg
            CharIsThere(i) = True
            frmChars.picChar(i).Visible = True
        Else
            CharIsThere(i) = False
            frmChars.picChar(i).Visible = False
            frmChars.lblCharDetails(i).Caption = "(blank)"
        End If
        
        Char_Sprite(i) = Val(Parse(n + 3))
        
        frmChars.picChar(i).Top = 36
        
        If CharIsThere(i) Then
            If Sprite_Size(Char_Sprite(i)).SizeY <> PIC_Y Then
                frmChars.picChar(i).Top = frmChars.picChar(i).Top - ((Sprite_Size(Char_Sprite(i)).SizeY - PIC_Y) / 2)
            End If
        End If
        
        DrawSelectedCharacter i, GameConfig.StandFrame
        
        n = n + 4
    Next
    
    frmChars.tmrChars.Interval = GameConfig.WalkAnim_Speed
    frmChars.tmrChars.Enabled = True
    
End Sub

' :::::::::::::::::::::::::::::::::
 ' :: Login was successful packet ::
 ' :::::::::::::::::::::::::::::::::
Private Sub HandleLoginOk(ByRef Parse() As String)
     ' Now we can receive game data
     MyIndex = CLng(Parse(1))
     
     frmChars.Visible = False
     
     SetStatus "Receiving game data..."
     
End Sub

' :::::::::::::::::::::::::::::::::::::::
 ' :: New character classes data packet ::
 ' :::::::::::::::::::::::::::::::::::::::
Private Sub HandleNewCharClasses(ByRef Parse() As String)
Dim n As Long
Dim i As Long
     
     n = 1
     
     ' Max classes
     Max_Classes = CByte(Parse(n))
     ReDim Class(1 To Max_Classes) As ClassRec
     
     n = n + 1
     
     For i = 1 To Max_Classes
         With Class(i)
             .Name = Parse(n)
             
             .Vital(Vitals.HP) = CLng(Parse(n + 1))
             .Vital(Vitals.MP) = CLng(Parse(n + 2))
             .Vital(Vitals.SP) = CLng(Parse(n + 3))
             
             .Stat(Stats.Strength) = CLng(Parse(n + 4))
             .Stat(Stats.Defense) = CLng(Parse(n + 5))
             .Stat(Stats.Speed) = CLng(Parse(n + 6))
             .Stat(Stats.Magic) = CLng(Parse(n + 7))
         End With
         
         n = n + 8
     Next
     
     ' Used for if the player is creating a new character
     frmNewChar.Caption = Game_Name & " [New Character]"
     frmNewChar.Visible = True
     frmStatus.Visible = False

     frmNewChar.cmbClass.Clear

     For i = 1 To Max_Classes
         frmNewChar.cmbClass.AddItem Trim$(Class(i).Name)
     Next
         

     frmNewChar.cmbClass.ListIndex = 0
     
     n = frmNewChar.cmbClass.ListIndex + 1
     
     frmNewChar.lblHP.Caption = CStr(Class(n).Vital(Vitals.HP))
     frmNewChar.lblMP.Caption = CStr(Class(n).Vital(Vitals.MP))
     frmNewChar.lblSP.Caption = CStr(Class(n).Vital(Vitals.SP))
 
     frmNewChar.lblStrength.Caption = CStr(Class(n).Stat(Stats.Strength))
     frmNewChar.lblDefense.Caption = CStr(Class(n).Stat(Stats.Defense))
     frmNewChar.lblSpeed.Caption = CStr(Class(n).Stat(Stats.Speed))
     frmNewChar.lblMagic.Caption = CStr(Class(n).Stat(Stats.Magic))
End Sub

' :::::::::::::::::::::::::
 ' :: Classes data packet ::
 ' :::::::::::::::::::::::::
Private Sub HandleClassesData(ByRef Parse() As String)
Dim n As Long
Dim i As Long
     
     n = 1
     
     ' Max classes
     Max_Classes = CByte(Parse(n))
     ReDim Preserve Class(1 To Max_Classes) As ClassRec
     
     n = n + 1
     
     For i = 1 To Max_Classes
         With Class(i)
             .Name = Parse(n)
             
             .Vital(Vitals.HP) = CLng(Parse(n + 1))
             .Vital(Vitals.MP) = CLng(Parse(n + 2))
             .Vital(Vitals.SP) = CLng(Parse(n + 3))
             
             .Stat(Stats.Strength) = CLng(Parse(n + 4))
             .Stat(Stats.Defense) = CLng(Parse(n + 5))
             .Stat(Stats.Speed) = CLng(Parse(n + 6))
             .Stat(Stats.Magic) = CLng(Parse(n + 7))
         End With
         
         n = n + 8
     Next
End Sub

' ::::::::::::::::::::
 ' :: In game packet ::
 ' ::::::::::::::::::::
Private Sub HandleInGame()
     InGame = True
     GameInit
     GameLoop
End Sub

' :::::::::::::::::::::::::::::
 ' :: Player inventory packet ::
 ' :::::::::::::::::::::::::::::
Private Sub HandlePlayerInv(ByRef Parse() As String)
Dim n As Long
Dim i As Long

     n = 1
     For i = 1 To MAX_INV
         SetPlayerInvItemNum MyIndex, i, CLng(Parse(n))
         SetPlayerInvItemValue MyIndex, i, CLng(Parse(n + 1))
         SetPlayerInvItemDur MyIndex, i, CLng(Parse(n + 2))
         
         n = n + 3
     Next
     UpdateInventory
 End Sub

' ::::::::::::::::::::::::::::::::::::
 ' :: Player inventory update packet ::
 ' ::::::::::::::::::::::::::::::::::::
Private Sub HandlePlayerInvUpdate(ByRef Parse() As String)
Dim n As Long

     n = CLng(Parse(1))
     
     SetPlayerInvItemNum MyIndex, n, CLng(Parse(2))
     SetPlayerInvItemValue MyIndex, n, CLng(Parse(3))
     SetPlayerInvItemDur MyIndex, n, CLng(Parse(4))
     UpdateInventory
End Sub

' ::::::::::::::::::::::::::::::::::
 ' :: Player worn equipment packet ::
 ' ::::::::::::::::::::::::::::::::::
Private Sub HandlePlayerWornEq(ByRef Parse() As String)
     SetPlayerEquipmentSlot MyIndex, CLng(Parse(1)), Armor
     SetPlayerEquipmentSlot MyIndex, CLng(Parse(2)), Weapon
     SetPlayerEquipmentSlot MyIndex, CLng(Parse(3)), Helmet
     SetPlayerEquipmentSlot MyIndex, CLng(Parse(4)), Shield
     UpdateInventory
End Sub

' ::::::::::::::::::::::
 ' :: Player hp packet ::
 ' ::::::::::::::::::::::
Private Sub HandlePlayerHp(ByRef Parse() As String)
Dim Index As Long

    If MyIndex < 1 Then Exit Sub
    
    Index = Val(Parse(3))
    
    Player(Index).MaxHP = CLng(Parse(1))
    SetPlayerVital Index, Vitals.HP, CLng(Parse(2))
    
    If Index = MyIndex Then
        If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
            frmMainGame.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
            
            If (GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP)) > 1 Then
                frmMainGame.picHP.Width = HPMPBar_Width
                Exit Sub
            End If
            
            frmMainGame.picHP.Width = HPMPBar_Width * (GetPlayerVital(MyIndex, Vitals.HP) / GetPlayerMaxVital(MyIndex, Vitals.HP))
        Else
            frmMainGame.picHP.Width = 0
        End If
    End If
    
End Sub

' ::::::::::::::::::::::
' :: Player mp packet ::
' ::::::::::::::::::::::
Private Sub HandlePlayerMp(ByRef Parse() As String)
    If MyIndex < 1 Then Exit Sub
    Player(MyIndex).MaxMP = CLng(Parse(1))
    SetPlayerVital MyIndex, Vitals.MP, CLng(Parse(2))
    If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
        frmMainGame.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
        
        If (GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP)) > 1 Then
            frmMainGame.picMP.Width = HPMPBar_Width
            Exit Sub
        End If
        
        frmMainGame.picMP.Width = HPMPBar_Width * (GetPlayerVital(MyIndex, Vitals.MP) / GetPlayerMaxVital(MyIndex, Vitals.MP))
    Else
        frmMainGame.picMP.Width = 0
    End If
End Sub

' ::::::::::::::::::::::
 ' :: Player sp packet ::
 ' ::::::::::::::::::::::
Private Sub HandlePlayerSp(ByRef Parse() As String)
    If MyIndex < 1 Then Exit Sub
    Player(MyIndex).MaxSP = CLng(Parse(1))
    SetPlayerVital MyIndex, Vitals.SP, CLng(Parse(2))
    'If GetPlayerMaxVital(MyIndex, Vitals.SP) > 0 Then
        'frmMainGame.lblSP.Caption = Int(GetPlayerVital(MyIndex, Vitals.SP) / GetPlayerMaxVital(MyIndex, Vitals.SP) * 100) & "%"
    '    frmMainGame.lblSP.Caption = GetPlayerVital(MyIndex, Vitals.SP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.SP)
    'End If
End Sub

' :::::::::::::::::::::::::
 ' :: Player stats packet ::
 ' :::::::::::::::::::::::::
Private Sub HandlePlayerStats(ByRef Parse() As String)
     SetPlayerStat MyIndex, Stats.Strength, CLng(Parse(1))
     SetPlayerStat MyIndex, Stats.Defense, CLng(Parse(2))
     SetPlayerStat MyIndex, Stats.Speed, CLng(Parse(3))
     SetPlayerStat MyIndex, Stats.Magic, CLng(Parse(4))
End Sub

Private Sub HandlePlayerStatBuffs(ByRef Parse() As String)
Dim StatName As String
Dim i As Long

    StatBuffed(Stats.Strength) = Val(Parse(1))
    StatBuffed(Stats.Defense) = Val(Parse(2))
    StatBuffed(Stats.Speed) = Val(Parse(3))
    StatBuffed(Stats.Magic) = Val(Parse(4))
    
    With frmMainGame
        For i = 1 To UBound(Player(MyIndex).Stat)
            Select Case i
                Case 1
                    StatName = "Strength"
                Case 2
                    StatName = "Defense"
                Case 3
                    StatName = "Speed"
                Case 4
                    StatName = "Magic"
            End Select
            .lblStat(i - 1).Caption = StatName & ": (" & StatBuffed(i) & "/" & Player(MyIndex).Stat(i) & ")"
        Next
    End With
    
End Sub

 ' :::::::::::::::::::::::::
 ' :: Player data packet ::
 ' ::::::::::::::::::::::::
Private Sub HandlePlayerData(ByRef Parse() As String)
Dim i As Long

     i = CLng(Parse(1))
     
     SetPlayerName i, Parse(2)
     SetPlayerSprite i, CLng(Parse(3))
     SetPlayerMap i, CLng(Parse(4))
     SetPlayerX i, CLng(Parse(5))
     SetPlayerY i, CLng(Parse(6))
     SetPlayerDir i, CLng(Parse(7))
     SetPlayerAccess i, CLng(Parse(8))
     SetPlayerPK i, CLng(Parse(9))
     
     ' Check if the player is the client player, and if so reset directions
     If i = MyIndex Then
         DirUp = False
         DirDown = False
         DirLeft = False
         DirRight = False
     End If
     
     ' Make sure they aren't walking
     Player(i).Moving = 0
     Player(i).XOffset = 0
     Player(i).YOffset = 0
     
End Sub

' ::::::::::::::::::::::::::::
 ' :: Player movement packet ::
 ' ::::::::::::::::::::::::::::
Private Sub HandlePlayerMove(ByRef Parse() As String)
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim n As Byte

     i = CLng(Parse(1))
     X = CLng(Parse(2))
     Y = CLng(Parse(3))
     Dir = CLng(Parse(4))
     n = CByte(Parse(5))

     SetPlayerX i, X
     SetPlayerY i, Y
     SetPlayerDir i, Dir
             
     Player(i).XOffset = 0
     Player(i).YOffset = 0
     Player(i).Moving = n
     
     Select Case GetPlayerDir(i)
         Case E_Direction.Up_
             Player(i).YOffset = PIC_Y
         Case E_Direction.Down_
             Player(i).YOffset = -PIC_Y
         Case E_Direction.Left_
             Player(i).XOffset = PIC_X
         Case E_Direction.Right_
             Player(i).XOffset = -PIC_X
     End Select
End Sub

' :::::::::::::::::::::::::
 ' :: Npc movement packet ::
 ' :::::::::::::::::::::::::
Private Sub HandleNpcMove(ByRef Parse() As String)

    If InEditor Then Exit Sub
    
Dim i As Long
Dim X As Long
Dim Y As Long
Dim Dir As Long
Dim n As Byte

     i = CLng(Parse(1))
     X = CLng(Parse(2))
     Y = CLng(Parse(3))
     Dir = CLng(Parse(4))
     n = CByte(Parse(5))

     MapNpc(i).X = X
     MapNpc(i).Y = Y
     MapNpc(i).Dir = Dir
     MapNpc(i).XOffset = 0
     MapNpc(i).YOffset = 0
     MapNpc(i).Moving = n
     
     Select Case GetMapNpcDir(i)
         Case E_Direction.Up_
             MapNpc(i).YOffset = PIC_Y
         Case E_Direction.Down_
             MapNpc(i).YOffset = -PIC_Y
         Case E_Direction.Left_
             MapNpc(i).XOffset = PIC_X
         Case E_Direction.Right_
             MapNpc(i).XOffset = -PIC_X
     End Select
End Sub

' :::::::::::::::::::::::::::::
 ' :: Player direction packet ::
 ' :::::::::::::::::::::::::::::
Private Sub HandlePlayerDir(ByRef Parse() As String)
Dim i As Long
Dim Dir As Byte

     i = CLng(Parse(1))
     Dir = CByte(Parse(2))
     SetPlayerDir i, Dir
     
     Player(i).XOffset = 0
     Player(i).YOffset = 0
     Player(i).Moving = 0
End Sub

' ::::::::::::::::::::::::::
 ' :: NPC direction packet ::
 ' ::::::::::::::::::::::::::
Private Sub HandleNpcDir(ByRef Parse() As String)
Dim i As Long
Dim Dir As Byte

     i = CLng(Parse(1))
     Dir = CByte(Parse(2))
     MapNpc(i).Dir = Dir
     
     MapNpc(i).XOffset = 0
     MapNpc(i).YOffset = 0
     MapNpc(i).Moving = 0
End Sub

' :::::::::::::::::::::::::::::::
 ' :: Player XY location packet ::
 ' :::::::::::::::::::::::::::::::
Private Sub HandlePlayerXY(ByRef Parse() As String)
Dim X As Long
Dim Y As Long

     X = CLng(Parse(1))
     Y = CLng(Parse(2))
     
     SetPlayerX MyIndex, X
     SetPlayerY MyIndex, Y
     
     ' Make sure they aren't walking
     Player(MyIndex).Moving = 0
     Player(MyIndex).XOffset = 0
     Player(MyIndex).YOffset = 0
End Sub

' ::::::::::::::::::::::::::
 ' :: Player attack packet ::
 ' ::::::::::::::::::::::::::
Private Sub HandleAttack(ByRef Parse() As String)
Dim i As Long

     i = CLng(Parse(1))
     
     ' Set player to attacking
     Player(i).Attacking = 1
     Player(i).AttackTimer = GetTickCountNew
End Sub

' :::::::::::::::::::::::
 ' :: NPC attack packet ::
 ' :::::::::::::::::::::::
Private Sub HandleNpcAttack(ByRef Parse() As String)
Dim i As Long

     i = CLng(Parse(1))
     
     ' Set player to attacking
     MapNpc(i).Attacking = 1
     MapNpc(i).AttackTimer = GetTickCountNew
End Sub

' ::::::::::::::::::::::::::
 ' :: Check for map packet ::
 ' ::::::::::::::::::::::::::
Private Sub HandleCheckForMap(ByRef Parse() As String)
Dim X As Long
Dim Y As Long
Dim i As Long
    
    ' Erase all players except self
    For i = 1 To MAX_PLAYERS
        If i <> MyIndex Then SetPlayerMap i, 0
    Next
    
    ' Erase all temporary tile values
    ClearTempTile
    
    ' Get map num
    X = CLng(Parse(1))
    
    ' Get revision
    Y = CLng(Parse(2))
    
    If FileExist(MAP_PATH & "map" & X & MAP_EXT) Then
        LoadMap X
        
        ' Check to see if the revisions match
        If Map.Revision = Y Then
            ' We do so we dont need the map
            SendData CNeedMap & SEP_CHAR & 0 & END_CHAR
            Exit Sub
        End If
    End If
    
    ' Either the revisions didn't match or we dont have the map, so we need it
    SendData CNeedMap & SEP_CHAR & 1 & END_CHAR
    
End Sub

' :::::::::::::::::::::
 ' :: Map data packet ::
 ' :::::::::::::::::::::
Private Sub HandleMapData(ByRef Parse() As String)
Dim n As Long
Dim X As Long
Dim Y As Long
Dim i As Long

    n = 1
    
    Map.Name = Parse(n + 1)
    Map.Revision = CLng(Parse(n + 2))
    Map.Moral = CByte(Parse(n + 3))
    Map.Up = CInt(Parse(n + 4))
    Map.Down = CInt(Parse(n + 5))
    Map.Left = CInt(Parse(n + 6))
    Map.Right = CInt(Parse(n + 7))
    Map.Music = Parse(n + 8)
    Map.BootMap = Val(Parse(n + 9))
    Map.BootX = Val(Parse(n + 10))
    Map.BootY = Val(Parse(n + 11))
    
    n = n + 12
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
        
            For i = 0 To UBound(Map.Tile(X, Y).Layer)
                Map.Tile(X, Y).Layer(i) = Val(Parse(n))
                Map.Tile(X, Y).LayerSet(i) = Val(Parse(n + 1))
                n = n + 2
            Next
            
            Map.Tile(X, Y).Type = Val(Parse(n))
            Map.Tile(X, Y).Data1 = Val(Parse(n + 1))
            Map.Tile(X, Y).Data2 = Val(Parse(n + 2))
            Map.Tile(X, Y).Data3 = Val(Parse(n + 3))
            
            n = n + 4
            
        Next
    Next
    
    MAX_MAP_NPCS = Val(Parse(n))
    
    If MAX_MAP_NPCS <> 10 Then MAX_MAP_NPCS = 10
    
    ReDim MapSpawn.Npc(1 To MAX_MAP_NPCS)
    
    On Error Resume Next
    If UBound(MapNpc) <> MAX_MAP_NPCS Then ReDim MapNpc(1 To MAX_MAP_NPCS)
    
    n = n + 1
    
    For X = 1 To MAX_MAP_NPCS
        MapSpawn.Npc(X).Num = CLng(Parse(n))
        MapSpawn.Npc(X).X = CLng(Parse(n + 1))
        MapSpawn.Npc(X).Y = CLng(Parse(n + 2))
        n = n + 3
    Next
    
    ' Save the map
    SaveMap CLng(Parse(1))
     
    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InEditor Then
        InEditor = False
        frmMainGame.picMapEditor.Visible = False
        frmMainGame.Width = Default_MenuWidth
        If frmAttrib.Visible Then Unload frmAttrib: ClearMapAttribs
        If frmMapProperties.Visible Then Unload frmMapProperties
    End If
    
End Sub

' :::::::::::::::::::::::::::
 ' :: Map items data packet ::
 ' :::::::::::::::::::::::::::
Private Sub HandleMapItemData(ByRef Parse() As String)
Dim n As Long
Dim i As Long

     n = 1
     
     For i = 1 To MAX_MAP_ITEMS
         MapItem(i).Num = CByte(Parse(n))
         MapItem(i).Value = CLng(Parse(n + 1))
         MapItem(i).Dur = CInt(Parse(n + 2))
         MapItem(i).X = CByte(Parse(n + 3))
         MapItem(i).Y = CByte(Parse(n + 4))
         MapItem(i).Anim = CLng(Parse(n + 5))
         
         n = n + 6
     Next
     
End Sub

' :::::::::::::::::::::::::
 ' :: Map npc data packet ::
 ' :::::::::::::::::::::::::
Private Sub HandleMapNpcData(ByRef Parse() As String)
Dim n As Long
Dim i As Long

    n = 2
    
    For i = 1 To UBound(MapNpc)
        MapNpc(i).Num = CByte(Parse(n))
        MapNpc(i).X = CByte(Parse(n + 1))
        MapNpc(i).Y = CByte(Parse(n + 2))
        MapNpc(i).Dir = CByte(Parse(n + 3))
        
        n = n + 4
    Next
    
End Sub

' :::::::::::::::::::::::::::::::
 ' :: Map send completed packet ::
 ' :::::::::::::::::::::::::::::::
Private Sub HandleMapDone()
Dim i As Long
Dim X As Long
Dim Y As Long
Dim TilesetLoaded() As Boolean

    GettingMap = False
    
    ' Play music
    If LenB(Trim$(Map.Music)) > 0 Then
        DirectMusic_StopMidi
        DirectMusic_PlayMidi Trim$(Map.Music) & MUSIC_EXT
    Else
        DirectMusic_StopMidi
    End If
    
    UpdateDrawMapName
    
    ReDim TilesetLoaded(0 To MAX_TILESETS)
    
    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            For i = 0 To UBound(Map.Tile(X, Y).Layer)
                If Not TilesetLoaded(Map.Tile(X, Y).LayerSet(i)) Then
                    InitTileSurf Map.Tile(X, Y).LayerSet(i)
                    TilesetLoaded(Map.Tile(X, Y).LayerSet(i)) = True
                End If
            Next
        Next
    Next
    
    CanMoveNow = True
    
End Sub

 ' ::::::::::::::::::::
 ' :: Social packets ::
 ' ::::::::::::::::::::
 
Private Sub HandleMessage(ByRef Parse() As String)

    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text)
    frmMainGame.txtChat.SelColor = ColorTable(Color.Black)
    frmMainGame.txtChat.SelText = vbNewLine & Left$(Parse(1), 1)
    
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text)
    frmMainGame.txtChat.SelColor = ColorTable(CLng(Parse(5)))
    frmMainGame.txtChat.SelText = Mid$(Parse(1), 2, Len(Parse(1)) - 2) 'Left$(Parse(1), 1)
    
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text)
    frmMainGame.txtChat.SelColor = ColorTable(Color.Black)
    frmMainGame.txtChat.SelText = Right$(Parse(1), 1) & " "
    
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text)
    frmMainGame.txtChat.SelColor = ColorTable(CLng(Parse(2)))
    frmMainGame.txtChat.SelText = Parse(3)
    
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text)
    frmMainGame.txtChat.SelColor = ColorTable(Color.Black)
    frmMainGame.txtChat.SelText = ": "
    
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text)
    frmMainGame.txtChat.SelColor = ColorTable(Color.Grey) 'ColorTable(CLng(Parse(5)))
    frmMainGame.txtChat.SelText = Parse(4)
    
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text) - 1
    
    ' Prevent players from name spoofing
    frmMainGame.txtChat.SelHangingIndent = 10
    
End Sub

Private Sub HandleGlobalMsg(ByRef Parse() As String)
     AddText Parse(1), CInt(Parse(2))
End Sub

Private Sub HandlePlayerMsg(ByRef Parse() As String)
     AddText Parse(1), CInt(Parse(2))
End Sub

Private Sub HandleMapMsg(ByRef Parse() As String)
     AddText Parse(1), CInt(Parse(2))
End Sub

Private Sub HandleAdminMsg(ByRef Parse() As String)
     AddText Parse(1), CInt(Parse(2))
End Sub

' ::::::::::::::::::::::::
 ' :: Refresh editor packet ::
 ' ::::::::::::::::::::::::
Private Sub HandleRefresh()
    Dim i As Long
    
    frmIndex.lstIndex.Clear
    
    Select Case Editor
        Case GameEditor.Item_
            For i = 1 To MAX_ITEMS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
            Next
        Case GameEditor.NPC_
            For i = 1 To MAX_NPCS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
            Next
        Case GameEditor.Shop_
            For i = 1 To MAX_SHOPS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
            Next
        Case GameEditor.Spell_
            For i = 1 To MAX_SPELLS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
            Next
        Case GameEditor.Sign_
            For i = 1 To MAX_SIGNS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Sign(i).Name)
            Next
        Case GameEditor.Anim_
            For i = 1 To MAX_ANIMS
                frmIndex.lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
            Next
    End Select
     
    frmIndex.lstIndex.ListIndex = 0
     
End Sub

' :::::::::::::::::::::::
 ' :: Item spawn packet ::
 ' :::::::::::::::::::::::
Private Sub HandleSpawnItem(ByRef Parse() As String)
Dim n As Long

     n = CLng(Parse(1))
     
     MapItem(n).Num = CByte(Parse(2))
     MapItem(n).Value = CLng(Parse(3))
     MapItem(n).Dur = CInt(Parse(4))
     MapItem(n).X = CByte(Parse(5))
     MapItem(n).Y = CByte(Parse(6))
     MapItem(n).Anim = CLng(Parse(7))
     MapItem(n).AnimItem = CLng(Parse(8))
     
End Sub

' ::::::::::::::::::::::::
 ' :: Item editor packet ::
 ' ::::::::::::::::::::::::
Private Sub HandleItemEditor()
Dim i As Long

     Editor = 1

     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_ITEMS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Item(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
End Sub

' ::::::::::::::::::::::::
 ' :: Update item packet ::
 ' ::::::::::::::::::::::::
Private Sub HandleUpdateItem(ByRef Parse() As String)
Dim n As Long
Dim i As Long
Dim ParseCount As Long

    n = CLng(Parse(1))
    
    ' Update the item
    Item(n).Name = Parse(2)
    Item(n).Pic = CInt(Parse(3))
    Item(n).Type = CByte(Parse(4))
    Item(n).Anim = CLng(Parse(5))
    Item(n).CostItem = CLng(Parse(6))
    Item(n).CostAmount = CLng(Parse(7))
    Item(n).Data1 = CLng(Parse(8))
    Item(n).Data2 = CLng(Parse(9))
    Item(n).Data3 = CLng(Parse(10))
    
    ParseCount = 11
    
    For i = 1 To Stats.Stat_Count - 1
        Item(n).BuffStats(i) = Val(Parse(ParseCount))
        ParseCount = ParseCount + 1
    Next
    
    For i = 1 To Vitals.Vital_Count - 1
        Item(n).BuffVitals(i) = Val(Parse(ParseCount))
        ParseCount = ParseCount + 1
    Next
    
    For i = 1 To Item_Requires.Count - 1
        Item(n).Required(i) = Val(Parse(ParseCount))
        ParseCount = ParseCount + 1
    Next
    
    UpdateInventory
    
End Sub

' ::::::::::::::::::::::
 ' :: Edit item packet ::
 ' ::::::::::::::::::::::
Private Sub HandleEditItem(ByRef Parse() As String)
Dim n As Long
Dim LoopI As Long
Dim Packetcount As Long

     n = CLng(Parse(1))
     
     ' Update the item
     Item(n).Name = Parse(2)
     Item(n).Pic = CInt(Parse(3))
     Item(n).Type = CByte(Parse(4))
     Item(n).Durability = CInt(Parse(5))
     Item(n).Anim = CLng(Parse(6))
     Item(n).CostItem = CLng(Parse(7))
     Item(n).CostAmount = CLng(Parse(8))
     
     Packetcount = 9
     
     For LoopI = 1 To Stats.Stat_Count - 1
        Item(n).BuffStats(LoopI) = Val(Parse(Packetcount))
        Packetcount = Packetcount + 1
     Next
     
     For LoopI = 1 To Vitals.Vital_Count - 1
        Item(n).BuffVitals(LoopI) = Val(Parse(Packetcount))
        Packetcount = Packetcount + 1
     Next
     
     For LoopI = 0 To Item_Requires.Count - 1
        Item(n).Required(LoopI) = Val(Parse(Packetcount))
        Packetcount = Packetcount + 1
     Next
     
     ' Initialize the item editor
     ItemEditorInit
     
End Sub

' ::::::::::::::::::::::
 ' :: Npc spawn packet ::
 ' ::::::::::::::::::::::
Private Sub HandleSpawnNpc(ByRef Parse() As String)
Dim n As Long

     n = CLng(Parse(1))
     
     MapNpc(n).Num = CByte(Parse(2))
     MapNpc(n).X = CByte(Parse(3))
     MapNpc(n).Y = CByte(Parse(4))
     MapNpc(n).Dir = CByte(Parse(5))
     
     ' Client use only
     MapNpc(n).XOffset = 0
     MapNpc(n).YOffset = 0
     MapNpc(n).Moving = 0
End Sub

' :::::::::::::::::::::
 ' :: Npc dead packet ::
 ' :::::::::::::::::::::
 Private Sub HandleNpcDead(ByRef Parse() As String)
 Dim n As Long
 
     n = CLng(Parse(1))
     
     MapNpc(n).Num = 0
     MapNpc(n).X = 0
     MapNpc(n).Y = 0
     MapNpc(n).Dir = 0
     
     ' Client use only
     MapNpc(n).XOffset = 0
     MapNpc(n).YOffset = 0
     MapNpc(n).Moving = 0
End Sub

' :::::::::::::::::::::::
 ' :: Npc editor packet ::
 ' :::::::::::::::::::::::
Private Sub HandleNpcEditor()
Dim i As Long

    Editor = 2
     
     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_NPCS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Npc(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
End Sub

' :::::::::::::::::::::::
 ' :: Update npc packet ::
 ' :::::::::::::::::::::::
Private Sub HandleUpdateNpc(ByRef Parse() As String)
Dim n As Long

     n = CLng(Parse(1))
     
     ' Update the item
     Npc(n).Name = Parse(2)
     Npc(n).AttackSay = vbNullString
     Npc(n).Sprite = CInt(Parse(3))
     Npc(n).HP = CLng(Parse(4))
     Npc(n).SpawnSecs = 0
     Npc(n).Behavior = 0
     Npc(n).Range = 0
     Npc(n).DropChance = 0
     Npc(n).DropItem = 0
     Npc(n).DropItemValue = 0
     Npc(n).Stat(Stats.Strength) = 0
     Npc(n).Stat(Stats.Defense) = 0
     Npc(n).Stat(Stats.Speed) = 0
     Npc(n).Stat(Stats.Magic) = 0
End Sub

' :::::::::::::::::::::
 ' :: Edit npc packet ::
 ' :::::::::::::::::::::
Private Sub HandleEditNpc(ByRef Parse() As String)
Dim n As Long
Dim i As Long
Dim ii As Long

     n = CLng(Parse(1))
     
     ' Update the npc
     Npc(n).Name = Parse(2)
     Npc(n).AttackSay = Parse(3)
     Npc(n).Sprite = CInt(Parse(4))
     Npc(n).SpawnSecs = CLng(Parse(5))
     Npc(n).Behavior = CByte(Parse(6))
     Npc(n).Range = CByte(Parse(7))
     Npc(n).DropChance = CInt(Parse(8))
     Npc(n).DropItem = CByte(Parse(9))
     Npc(n).DropItemValue = CInt(Parse(10))
     Npc(n).Stat(Stats.Strength) = CByte(Parse(11))
     Npc(n).Stat(Stats.Defense) = CByte(Parse(12))
     Npc(n).Stat(Stats.Speed) = CByte(Parse(13))
     Npc(n).Stat(Stats.Magic) = CByte(Parse(14))
     Npc(n).GivesGuild = CByte(Parse(15))
     
     ii = 16
     
     For i = 0 To UBound(Npc(n).Sound)
        Npc(n).Sound(i) = Parse(ii)
        ii = ii + 1
     Next
     
     For i = 0 To UBound(Npc(n).Reflection)
        Npc(n).Reflection(i) = Val(Parse(ii))
        ii = ii + 1
     Next
     
     ' Initialize the npc editor
     NpcEditorInit
     
End Sub

' ::::::::::::::::::::
 ' :: Map key packet ::
 ' ::::::::::::::::::::
Private Sub HandleMapKey(ByRef Parse() As String)
Dim n As Long
Dim X As Long
Dim Y As Long

     X = CLng(Parse(1))
     Y = CLng(Parse(2))
     n = CLng(Parse(3))
     
     TempTile(X, Y).DoorOpen = n
End Sub

' :::::::::::::::::::::
 ' :: Edit map packet ::
 ' :::::::::::::::::::::
Private Sub HandleEditMap()
     MapEditorInit
End Sub

' ::::::::::::::::::::::::
 ' :: Shop editor packet ::
 ' ::::::::::::::::::::::::
Private Sub HandleShopEditor()
Dim i As Long

     Editor = 4
     
     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_SHOPS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Shop(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
End Sub

' ::::::::::::::::::::::::
 ' :: Update shop packet ::
 ' ::::::::::::::::::::::::
Private Sub HandleUpdateShop(ByRef Parse() As String)
Dim n As Long

     n = CLng(Parse(1))
     
     ' Update the shop name
     Shop(n).Name = Parse(2)
End Sub

' ::::::::::::::::::::::
 ' :: Edit shop packet ::
 ' ::::::::::::::::::::::
Private Sub HandleEditShop(ByRef Parse() As String)
Dim n As Long
Dim i As Long
Dim ShopNum As Long
Dim GiveItem As Long
Dim GiveValue As Long
Dim GetItem As Long
Dim GetValue As Long

     ShopNum = CLng(Parse(1))
     
     ' Update the shop
     Shop(ShopNum).Name = Parse(2)
     Shop(ShopNum).JoinSay = Parse(3)
     Shop(ShopNum).LeaveSay = Parse(4)
     Shop(ShopNum).FixesItems = CByte(Parse(5))
     
     n = 6
     
     For i = 1 To MAX_TRADES
         
         GiveItem = CLng(Parse(n))
         GiveValue = CLng(Parse(n + 1))
         GetItem = CLng(Parse(n + 2))
         GetValue = CLng(Parse(n + 3))
         
         Shop(ShopNum).TradeItem(i).GiveItem = GiveItem
         Shop(ShopNum).TradeItem(i).GiveValue = GiveValue
         Shop(ShopNum).TradeItem(i).GetItem = GetItem
         Shop(ShopNum).TradeItem(i).GetValue = GetValue
         
         n = n + 4
     Next
     
     ' Initialize the shop editor
     ShopEditorInit
     
End Sub

' :::::::::::::::::::::::::
' :: Anim editor packet  ::
' :::::::::::::::::::::::::
Private Sub HandleAnimEditor()
Dim i As Long

     Editor = GameEditor.Anim_
     
     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_ANIMS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Animation(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
     
End Sub
     

' :::::::::::::::::::::::::
' :: Sign editor packet  ::
' :::::::::::::::::::::::::
Private Sub HandleSignEditor()
Dim i As Long

     Editor = GameEditor.Sign_
     
     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_SIGNS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Sign(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
     
End Sub

' :::::::::::::::::::::::::
 ' :: Spell editor packet ::
 ' :::::::::::::::::::::::::
Private Sub HandleSpellEditor()
Dim i As Long

     Editor = 3
     
     frmIndex.lstIndex.Clear
     
     ' Add the names
     For i = 1 To MAX_SPELLS
         frmIndex.lstIndex.AddItem i & ": " & Trim$(Spell(i).Name)
     Next
     
     frmIndex.lstIndex.ListIndex = 0
     frmIndex.Show vbModal
End Sub

' ::::::::::::::::::::::::
 ' :: Update spell packet ::
 ' ::::::::::::::::::::::::
Private Sub HandleUpdateSpell(ByRef Parse() As String)
Dim n As Long

     n = CLng(Parse(1))
     
     ' Update the spell name
     Spell(n).Name = Parse(2)
     Spell(n).Anim = CLng(Parse(3))
     Spell(n).Icon = CInt(Parse(4))
     Spell(n).Timer = CLng(Parse(5))
     Spell(n).Data1 = CLng(Parse(6))
     Spell(n).AOE = CLng(Parse(7))
     
     DrawSpellList
     
End Sub

' ::::::::::::::::::::::::
 ' :: Update anim packet ::
 ' ::::::::::::::::::::::::
Private Sub HandleUpdateAnim(ByRef Parse() As String)
Dim n As Long

     n = CLng(Parse(1))
     
     ' Update the spell name
     Animation(n).Name = Parse(2)
     Animation(n).Height = Val(Parse(3))
     Animation(n).Width = Val(Parse(4))
     Animation(n).Pic = Val(Parse(5))
     Animation(n).Delay = Val(Parse(6))
     
End Sub

' :::::::::::::::::::::::
 ' :: Edit anim packet ::
 ' ::::::::::::::::::::::
Private Sub HandleEditAnim(ByRef Parse() As String)
Dim n As Long
Dim LoopI As Long
Dim Packetcount As Long

    n = CLng(Parse(1))
    
    ' Update the sign
    Animation(n).Name = Parse(2)
    Animation(n).Delay = Val(Parse(3))
    Animation(n).Height = Val(Parse(4))
    Animation(n).Width = Val(Parse(5))
    
    ' Initialize the sign editor
    AnimEditorInit
    
End Sub

' ::::::::::::::::::::::::
 ' :: Update sign packet ::
 ' ::::::::::::::::::::::::
Private Sub HandleUpdateSign(ByRef Parse() As String)
Dim n As Long

     n = CLng(Parse(1))
     
     ' Update the spell name
     Sign(n).Name = Parse(2)
     
End Sub

' :::::::::::::::::::::::
 ' :: Edit sign packet ::
 ' ::::::::::::::::::::::
Private Sub HandleEditSign(ByRef Parse() As String)
Dim n As Long
Dim LoopI As Long
Dim Packetcount As Long

    n = CLng(Parse(1))
    
    ' Update the sign
    Sign(n).Name = Parse(2)
    
    ReDim Preserve Sign(n).Section(0 To Val(Parse(3)))
    ReDim Preserve SignSection(0 To Val(Parse(3)))
    
    Packetcount = 4
    
    For LoopI = 0 To UBound(Sign(n).Section)
        Sign(n).Section(LoopI) = Parse(Packetcount)
        Packetcount = Packetcount + 1
    Next
    
    ' Initialize the sign editor
    SignEditorInit
    
End Sub

' :::::::::::::::::::::::
 ' :: Edit spell packet ::
Private Sub HandleEditSpell(ByRef Parse() As String)
Dim n As Long

    n = CLng(Parse(1))
    
    ' Update the spell
    With Spell(n)
        .Name = Parse(2)
        .MPReq = CByte(Parse(3))
        .Type = CByte(Parse(4))
        .Anim = CByte(Parse(5))
        .Icon = CInt(Parse(6))
        .Range = CByte(Parse(7))
        .Data1 = CInt(Parse(8))
        .Data2 = CInt(Parse(9))
        .Data3 = CInt(Parse(10))
        .CastSound = Parse(11)
        .AOE = Val(Parse(12))
        .Timer = Val(Parse(13))
    End With
    
    ' Initialize the spell editor
    SpellEditorInit
    
End Sub

' ::::::::::::::::::
 ' :: Trade packet ::
 ' ::::::::::::::::::
 Private Sub HandleTrade(ByRef Parse() As String)
 Dim n As Long
 Dim i As Long
 Dim ShopNum As Long
 Dim GiveItem As Long
 Dim GiveValue As Long
 Dim GetItem As Long
 Dim GetValue As Long
 
    If Not frmMainGame.picInv.Visible Then frmMainGame.picInventory_Click
    
    ShopNum = CLng(Parse(1))
    
    If CByte(Parse(2)) = 1 Then
        'frmTrade.lblFixItem.Visible = True
    Else
        'frmTrade.lblFixItem.Visible = False
    End If
    
    If LenB(Parse(3)) > 0 Then frmMainGame.lblWelcome.Caption = vbQuote & Parse(3) & vbQuote
    
    n = 4
    
    For i = 1 To MAX_TRADES
        GiveItem = CLng(Parse(n))
        GiveValue = CLng(Parse(n + 1))
        GetItem = CLng(Parse(n + 2))
        GetValue = CLng(Parse(n + 3))
        
        ShopTrade.TradeItem(i).GiveItem = GiveItem
        ShopTrade.TradeItem(i).GiveValue = GiveValue
        ShopTrade.TradeItem(i).GetItem = GetItem
        ShopTrade.TradeItem(i).GetValue = GetValue
        
        n = n + 4
    Next
    
    frmMainGame.picShop.Visible = True
    
    frmMainGame.shpSelect.Top = 1
    frmMainGame.shpSelect.Left = 1
    frmMainGame.shpSelect.Visible = False
    
    DrawShopList
    
End Sub

' :::::::::::::::::::
 ' :: Spells packet ::
 ' :::::::::::::::::::
Private Sub HandleSpells(ByRef Parse() As String)
Dim i As Long

    ' Put spells known in player record
    For i = 1 To MAX_PLAYER_SPELLS
        Player(MyIndex).Spell(i) = CByte(Parse(i))
    Next
    
    DrawSpellList
     
End Sub

Private Sub HandleLeft(ByRef Parse() As String)
    ClearPlayer Val(Parse(1))
End Sub

Private Sub HandleConfigPass(ByRef Parse() As String)

    If Val(Parse(1)) = 1 Then
        Password_Confirmed = True
        CheckedStuff = True
    Else
        Password_Confirmed = True
        MsgBox "Invalid " & CONFIG_FILE & " file. (wrong password)", , "Error"
        DestroyGame
    End If
    
End Sub

Private Sub HandleGameOptions(ByRef Parse() As String)
Dim i As Long
Dim ii As Byte

    With GameConfig
        Game_Name = Parse(1)
        GAME_WEBSITE = Parse(2)
        .Sprite_Offset = Val(Parse(3))
        .Total_WalkFrames = Val(Parse(4))
        .Total_AttackFrames = Val(Parse(5))
        .WalkAnim_Speed = Val(Parse(6))
        Total_AnimFrames = Val(Parse(7))
        .StandFrame = Val(Parse(8)) - 1
        Direction_Anim(0) = Val(Parse(9)) - 1
        Direction_Anim(1) = Val(Parse(10)) - 1
        Direction_Anim(2) = Val(Parse(11)) - 1
        Direction_Anim(3) = Val(Parse(12)) - 1
        MAX_PLAYERS = CLng(Parse(13))
        MAX_SHOPS = CLng(Parse(14))
        MAX_SPELLS = CLng(Parse(15))
        MAX_ITEMS = CLng(Parse(16))
        MAX_NPCS = CLng(Parse(17))
        MAX_MAPS = CLng(Parse(18))
        MAX_SIGNS = CLng(Parse(19))
        MAX_ANIMS = CLng(Parse(20))
        
        HandleNews Parse(21)
        
        If .Total_WalkFrames > 0 Then
            ii = 21
            ReDim .WalkFrame(1 To .Total_WalkFrames)
            For i = 1 To .Total_WalkFrames
                .WalkFrame(i) = Val(Parse(ii + 1))
                ii = ii + 1
            Next
        End If
        
        If .Total_AttackFrames > 0 Then
            ReDim .AttackFrame(1 To .Total_AttackFrames)
            For i = 1 To .Total_AttackFrames
                .AttackFrame(i) = Val(Parse(ii + 1))
                ii = ii + 1
            Next
        End If
        
        '1 = stand, then the walk frames, and finally attack frames
        Total_SpriteFrames = Total_AnimFrames \ 4
        
        Load_SpriteSizes
    End With
    
    ReDim Preserve Player(1 To MAX_PLAYERS) As PlayerRec
    ReDim Preserve Shop(1 To MAX_SHOPS) As ShopRec
    ReDim Preserve Spell(1 To MAX_SPELLS) As SpellRec
    ReDim Preserve Item(1 To MAX_ITEMS) As ItemRec
    ReDim Preserve Npc(1 To MAX_NPCS) As NpcRec
    ReDim Preserve Sign(1 To MAX_SIGNS) As SignRec
    ReDim Preserve Animation(1 To MAX_ANIMS) As AnimationRec
    
    Config_Received = True
    
    If frmMainMenu.Visible Then frmMainMenu.Caption = Game_Name & " [Main Menu]"
    
    CheckedTwice = True
    
End Sub

Private Sub HandleAnimation(ByRef Parse() As String)
Dim i As Long
Dim AnimNum As Long
Dim Victim As Long
Dim TargetType As Byte

    AnimNum = CLng(Parse(1))
    Victim = CLng(Parse(2))
    TargetType = CByte(Parse(3))
    
    For i = 1 To UBound(Animations)
        If Not Animations(i).Active Then
            Animations(i).Picture = Animation(AnimNum).Pic
            Animations(i).DelayTime = Animation(AnimNum).Delay
            Animations(i).Key = AnimNum
            
            If TargetType = E_Target.Player_ Then
                Animations(i).X = Player(Victim).X * PIC_X + Player(Victim).XOffset
                Animations(i).Y = Player(Victim).Y * PIC_Y + Player(Victim).YOffset
            ElseIf TargetType = E_Target.NPC_ Then
                Animations(i).X = MapNpc(Victim).X * PIC_X + MapNpc(Victim).XOffset
                Animations(i).Y = MapNpc(Victim).Y * PIC_Y + MapNpc(Victim).YOffset
            End If
            
            Animations(i).Timer = GetTickCountNew + Animation(AnimNum).Delay
            Animations(i).Active = True
            Exit For
            
        End If
    Next
    
End Sub

Private Sub HandleSoundPlay(ByRef Parse() As String)
    SoundPlay Parse(1) & SOUND_EXT
End Sub

Private Sub HandlePlayerPoints(ByRef Parse() As String)
    SetPlayerPOINTS MyIndex, Val(Parse(1))
End Sub

Private Sub HandleClassName(ByRef Parse() As String)
    CurrentClassName = Parse(1)
    frmMainGame.lblClassLevel.Caption = "Level " & GetPlayerLevel(MyIndex) & " " & Parse(1)
End Sub

Private Sub HandlePlayerLevel(ByRef Parse() As String)
    SetPlayerLevel MyIndex, Val(Parse(1))
End Sub

Private Sub HandleScrollingText(ByRef Parse() As String)

    Select Case UBound(Parse)
        Case 0
            ZeroMemory ByVal VarPtr(ScrollText), LenB(ScrollText)
            frmMainGame.lblSignText.Caption = vbNullString
            frmMainGame.picSign.Visible = False
            
        Case 3
            ScrollText.CurLetter = 1
            ScrollText.Text = Parse(1)
            ScrollText.KeyValue = Val(Parse(2))
            ScrollText.CurKey = Val(Parse(3))
            
            frmMainGame.lblPressEnter.Visible = False
            frmMainGame.lblSignText.Caption = vbNullString
            frmMainGame.picSign.Visible = True
            
            ScrollText.Running = True
            
    End Select
    
End Sub

Private Sub HandleGuildCreation()
    frmGuildCreation.Show vbModal
End Sub

Private Sub HandlePlayerGuild(ByRef Parse() As String)
Dim Index As Long

    Index = Val(Parse(1))
    
    If UBound(Parse) = 1 Then
        Player(Index).GuildName = vbNullString
        Player(Index).GuildRank = 0
        If Index = MyIndex Then frmMainGame.picGuildCP.Visible = False
        Exit Sub
    End If
    
    Player(Index).GuildName = Parse(2)
    Player(Index).GuildRank = Val(Parse(3))
    
    If Index = MyIndex Then
        frmMainGame.lblGuildName.Caption = "[" & Player(MyIndex).GuildName & "]"
        If Player(MyIndex).GuildRank = 4 Then
            frmMainGame.lblDisband.Caption = "disband"
            frmMainGame.lblPromote.Visible = True
            frmMainGame.lblDemote.Visible = True
            frmMainGame.txtInvite.Visible = True
            frmMainGame.lblInvite.Visible = True
            frmMainGame.lblKick.Visible = True
        Else
            frmMainGame.lblDisband.Caption = "leave"
            frmMainGame.lblPromote.Visible = False
            frmMainGame.lblDemote.Visible = False
            frmMainGame.txtInvite.Visible = False
            frmMainGame.lblInvite.Visible = False
            frmMainGame.lblKick.Visible = False
        End If
    End If
    
End Sub

Private Sub HandleGuildInvite(ByRef Parse() As String)
    SendData CInviteResponse & SEP_CHAR & (MsgBox("Do you wish to join the guild: " & Parse(1) & "?", vbYesNo, "Request") = vbYes) & SEP_CHAR & Val(Parse(2)) & END_CHAR
End Sub

Private Sub HandlePing()
    CurPing = GetTickCountNew - PingCounter
    PingCounter = 0
    WaitingonPing = False
End Sub

Private Sub HandleNpcHP(ByRef Parse() As String)

    MapNpc(Val(Parse(1))).Vital(Vitals.HP) = Val(Parse(2))
    
End Sub

Private Sub HandleNormalMsg(ByRef Parse() As String)

    MsgBox Parse(1), vbOKOnly, Game_Name
    
    If frmStatus.Visible Then frmStatus.Visible = False
    If CurrentWindow <> Val(Parse(2)) Then Windows(CurrentWindow).Visible = False
    
    CurrentWindow = Val(Parse(2))
    Windows(CurrentWindow).Visible = True
    
End Sub

Private Sub HandleCastSuccess(ByRef Parse() As String)

    If UBound(Parse) <> 2 Then
        Player(MyIndex).CastTimer(Val(Parse(1))) = GetTickCountNew + Spell(Player(MyIndex).Spell(Val(Parse(1)))).Timer
        Player(MyIndex).CastedSpell = YES
    Else
        Player(MyIndex).CastTimer(Val(Parse(1))) = 0
        Player(MyIndex).CastedSpell = NO
        frmMainGame.picSpellWaiting(Val(Parse(1))).Cls
    End If
    
End Sub

Private Sub HandleExpUpdate(ByRef Parse() As String)

    If MyIndex < 1 Then Exit Sub
    
    Player(MyIndex).Exp = Val(Parse(1))
    
    If GetPlayerExp(MyIndex) > 0 And Val(Parse(2)) > 0 Then
        frmMainGame.picTNL.Width = TNLBar_Width * (GetPlayerExp(MyIndex) / Val(Parse(2)))
    Else
        frmMainGame.picTNL.Width = 0
    End If
    
End Sub
