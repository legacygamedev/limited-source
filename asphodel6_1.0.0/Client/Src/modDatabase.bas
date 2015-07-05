Attribute VB_Name = "modDatabase"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

Public Function FileExist(ByVal FileName As String) As Boolean
    FileExist = LenB(Dir$(App.Path & "\" & FileName)) > 0
End Function

Public Sub AddLog(ByVal Text As String, Optional LogFile As String = LOG_PATH & LOG_DEBUG)
Dim FileName As String
Dim f As Long

    FileName = LogFile
    
    If Not FileExist(FileName) Then
        f = FreeFile
        Open App.Path & FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open App.Path & FileName For Append As #f
        Print #f, Time & ": " & Text
    Close #f
    
End Sub

Public Sub Load_Config()
Dim f As Long
Dim FileName As String

    FileName = CONFIG_FILE
    
    If Not FileExist(FileName) Then
        MsgBox "Error: Missing " & CONFIG_FILE & " file.", , "Error"
        DestroyGame
    End If
    
    FileName = App.Path & "\" & CONFIG_FILE
    
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Config
    Close #f
    
    GAME_IP = Encryption(DEFAULT_KEY, Trim$(Config.IP))
    GAME_PORT = Config.Port
    
End Sub

Public Sub Handle_TotalSprites()
Dim i As Long

    i = -1
    
    Do
    
        If Not FileExist(GFX_PATH & "Sprites\sprite" & i + 1 & GFX_EXT) Then Exit Do
        i = i + 1
    
        DoEvents
    Loop
    
    TOTAL_SPRITES = i
    
    If TOTAL_SPRITES < 0 Then
        MsgBox "No sprite images found.", , "Error"
        DestroyGame
    End If
    
    ReDim DDS_Sprite(0 To TOTAL_SPRITES)
    ReDim DDSD_Sprite(0 To TOTAL_SPRITES)
    ReDim Sprite_Size(0 To TOTAL_SPRITES)

End Sub

Public Sub Handle_TotalAnimGfx()
Dim i As Long

    i = -1
    
    Do
    
        If Not FileExist(GFX_PATH & "Anims\anim" & i + 1 & GFX_EXT) Then Exit Do
        i = i + 1
    
        DoEvents
    Loop
    
    TOTAL_ANIMGFX = i
    
    If TOTAL_ANIMGFX < 0 Then
        MsgBox "No animation images found.", , "Error"
        DestroyGame
    End If
    
    ReDim DDS_Anim(0 To TOTAL_ANIMGFX)
    ReDim DDSD_Anim(0 To TOTAL_ANIMGFX)
    
End Sub

Public Sub Load_GameConfig()
Dim FileName As String
Dim f As Long

    FileName = "info.ini"
    
    If Not FileExist(FileName) Then
    
        f = FreeFile
        Open FileName For Output As f
        Close f
        
        FileName = App.Path & "/" & FileName
        PutVar FileName, "BASIC", "Remember", CStr(0)
        PutVar FileName, "BASIC", "Sound_On", CStr(1)
        PutVar FileName, "BASIC", "Music_On", CStr(1)
        PutVar FileName, "BASIC", "ShowPlayerNames", CStr(1)
        PutVar FileName, "BASIC", "ShowNPCNames", CStr(1)
        PutVar FileName, "BASIC", "FPS_Cap", CStr(32)
        PutVar FileName, "BASIC", "Ping", CStr(0)
        Remember = False
        Sound_On = True
        Music_On = True
        ShowPNames = True
        ShowNNames = True
        FPS_Lock = 32
        PingEnabled = 0
        
    Else
    
        FileName = App.Path & "/" & FileName
        Remember = Val(GetVar(FileName, "BASIC", "Remember"))
        FPS_Lock = Val(GetVar(FileName, "BASIC", "FPS_Cap"))
        PingEnabled = Val(GetVar(FileName, "BASIC", "Ping"))
        
    End If
    
    If FPS_Lock < 1 Then
        FPS_Lock = 0
    Else
        If FPS_Lock < 32 Then FPS_Lock = 32
    End If
    
    If PingEnabled < 0 Or PingEnabled > 1 Then FPS_Lock = 0: PutVar FileName, "BASIC", "Ping", CStr(PingEnabled)
    
    PutVar FileName, "BASIC", "FPS_Cap", CStr(FPS_Lock)
    
    If FPS_Lock > 0 Then FPS_Lock = 1000 \ FPS_Lock - 1
    
    ' force either 0 or 1
    If Val(GetVar(FileName, "BASIC", "Sound_On")) >= 1 Then Sound_On = True Else Sound_On = False
    If Val(GetVar(FileName, "BASIC", "Music_On")) >= 1 Then Music_On = True Else Music_On = False
    If Val(GetVar(FileName, "BASIC", "ShowPlayerNames")) >= 1 Then ShowPNames = True Else ShowPNames = False
    If Val(GetVar(FileName, "BASIC", "ShowNPCNames")) >= 1 Then ShowNNames = True Else ShowNNames = False
    
End Sub

Public Sub Check_Password()

    SendData CConfigPass & SEP_CHAR & Trim$(Config.Password) & END_CHAR
    
    Do Until Password_Confirmed And Config_Received
        If Not IsConnected Then Exit Do
        DoEvents
        Sleep 1
    Loop
    
    If Not IsConnected Then
        MsgBox "The server appears to be down." & vbNewLine & _
               "Please check back later!", , "Error"
        'DestroyGame
        Exit Sub
    End If
    
End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space$(5000)
  
    GetPrivateProfileString$ Header, Var, szReturn, sSpaces, Len(sSpaces), File
  
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    WritePrivateProfileString$ Header, Var, Value, File
End Sub

Public Function Encryption(CodeKey As String, DataIn As String) As String
Dim lonDataPtr As Long
Dim strDataOut As String
Dim intXOrValue1 As Integer
Dim intXOrValue2 As Integer

    For lonDataPtr = 1 To Len(DataIn)
    
        intXOrValue1 = Asc(Mid$(DataIn, lonDataPtr, 1))
        intXOrValue2 = Asc(Mid$(CodeKey, ((lonDataPtr Mod Len(CodeKey)) + 1), 1))
        
        strDataOut = strDataOut + Chr$(intXOrValue1 Xor intXOrValue2)
    
    Next
    
    Encryption = strDataOut
   
End Function

Public Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
            
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map
    Close #f
End Sub

Public Sub LoadMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & MAP_PATH & "map" & MapNum & MAP_EXT
        
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , Map
    Close #f
End Sub

Public Sub ClearTempTile()
Dim X As Long
Dim Y As Long

    For X = 0 To MAX_MAPX
        For Y = 0 To MAX_MAPY
            TempTile(X, Y).DoorOpen = NO
        Next
    Next
End Sub

Public Sub ClearPlayer(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Player(Index)), LenB(Player(Index))
    Player(Index).Name = vbNullString
End Sub

Public Sub ClearItem(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(Item(Index)), LenB(Item(Index))
    Item(Index).Name = vbNullString
End Sub

Public Sub ClearItems()
Dim i As Long

    For i = 1 To MAX_ITEMS
        ClearItem (i)
    Next
End Sub

Public Sub ClearMapItem(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(MapItem(Index)), LenB(MapItem(Index))
End Sub

Public Sub ClearMap()
    ZeroMemory ByVal VarPtr(Map), LenB(Map)
    Map.Name = vbNullString
End Sub

Public Sub ClearMapItems()
Dim i As Long

    For i = 1 To MAX_MAP_ITEMS
        ClearMapItem i
    Next
End Sub

Public Sub ClearMapNpc(ByVal Index As Long)
    ZeroMemory ByVal VarPtr(MapNpc(Index)), LenB(MapNpc(Index))
End Sub

Public Sub ClearMapNpcs()
Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        ClearMapNpc i
    Next
End Sub

' **********************
' **   NPC functions  **
' **********************

Function GetMapNpcName(ByVal MapNpcNum As Long) As String
    If MapNpc(MapNpcNum).Num < 1 Then Exit Function
    GetMapNpcName = Trim$(Npc(MapNpc(MapNpcNum).Num).Name)
End Function

Function GetMapNpcSprite(ByVal MapNpcNum As Long) As Long
    If MapNpc(MapNpcNum).Num < 1 Then Exit Function
    GetMapNpcSprite = Npc(MapNpc(MapNpcNum).Num).Sprite
End Function

Function GetMapNpcDir(ByVal MapNpcNum As Long) As Byte
    If MapNpc(MapNpcNum).Num < 1 Then Exit Function
    GetMapNpcDir = MapNpc(MapNpcNum).Dir
End Function

Function GetMapNpcX(ByVal MapNpcNum As Long) As Byte
    If MapNpc(MapNpcNum).Num < 1 Then Exit Function
    GetMapNpcX = MapNpc(MapNpcNum).X
End Function

Function GetMapNpcY(ByVal MapNpcNum As Long) As Byte
    If MapNpc(MapNpcNum).Num < 1 Then Exit Function
    GetMapNpcY = MapNpc(MapNpcNum).Y
End Function

' **********************
' ** Player functions **
' **********************

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Name)
End Function

Public Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Name = Name
    If Index = MyIndex Then frmMainGame.lblPlayerName.Caption = "[ " & Name & " ]"
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Class
End Function

Public Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Sprite
End Function

Public Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Player(Index).Sprite = Sprite
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Level
End Function

Public Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    Player(Index).Level = Level
    frmMainGame.lblClassLevel.Caption = "Level " & Level & " " & CurrentClassName
End Sub

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Exp
End Function

Public Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Player(Index).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Access
End Function

Public Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Access = Access
    If frmAdmin.Visible Then
        If GetPlayerAccess(Index) < 1 Then
            Unload frmAdmin
        End If
    End If
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).PK
End Function

Public Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    GetPlayerVital = Player(Index).Vital(Vital)
End Function

Public Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Player(Index).Vital(Vital) = Value
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            GetPlayerMaxVital = Player(Index).MaxHP
        Case MP
            GetPlayerMaxVital = Player(Index).MaxMP
        Case SP
            GetPlayerMaxVital = Player(Index).MaxSP
    End Select
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    GetPlayerStat = Player(Index).Stat(Stat)
End Function

Public Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
Dim i As Long
Dim StatName As String

    Player(Index).Stat(Stat) = Value
    
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

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).POINTS
End Function

Public Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
Dim i As Long

    Player(Index).POINTS = POINTS
    
    For i = 0 To UBound(Player(MyIndex).Stat) - 1
        If POINTS > 0 Then
            frmMainGame.lblStatAdd(i).Visible = True
        Else
            frmMainGame.lblStatAdd(i).Visible = False
        End If
    Next
    
    frmMainGame.lblPoints.Caption = "Points: " & POINTS
    If POINTS > 0 Then frmMainGame.lblPoints.Visible = True Else frmMainGame.lblPoints.Visible = False
    
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    GetPlayerMap = Player(Index).Map
End Function

Public Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    Player(Index).Map = MapNum
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).X
End Function

Public Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Y
End Function

Public Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Dir
End Function

Public Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Dir = Dir
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemNum = Player(Index).Inv(InvSlot).Num
End Function

Public Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Inv(InvSlot).Value
End Function

Public Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).Dur
End Function

Public Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerEquipmentSlot(ByVal Index As Long, ByVal EquipmentSlot As Equipment) As Byte
    GetPlayerEquipmentSlot = Player(Index).Equipment(EquipmentSlot)
End Function

Public Sub SetPlayerEquipmentSlot(ByVal Index As Long, ByVal InvNum As Long, ByVal EquipmentSlot As Equipment)
    Player(Index).Equipment(EquipmentSlot) = InvNum
    If InvNum < 1 Then frmMainGame.picEquipment(EquipmentSlot).Cls: Exit Sub
    If Player(MyIndex).Inv(InvNum).Num > 0 Then Engine_BltToDC DDS_Item(Item(Player(MyIndex).Inv(InvNum).Num).Pic), Get_RECT, Get_RECT, frmMainGame.picEquipment(EquipmentSlot)
End Sub
