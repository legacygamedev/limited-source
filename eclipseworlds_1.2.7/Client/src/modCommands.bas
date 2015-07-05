Attribute VB_Name = "modCommands"
Option Explicit

Function GetPlayerName(ByVal Index As Long) As String
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Player(Index).Name)
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Name = Name
    
' Error handler
    Exit Sub
ErrorHandler:
    HandleError "SetPlayerName", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerGuild = Trim$(Player(Index).Guild)
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerGuild", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal GuildNum As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Player(Index).Guild = GuildNum
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerGuild", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Byte
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerGuildAccess = Player(Index).GuildAcc
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerGuildAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Access As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Player(Index).GuildAcc = Access
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerGuildAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Player(Index).Class
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Class = ClassNum
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerClass", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Player(Index).Sprite
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Sprite = Sprite
    Exit Sub
     
' Error handler
ErrorHandler:
    HandleError "SetPlayerSprite", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Player(Index).Level
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Level = Level
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(Index) + 1) ^ 3 - (6 * (GetPlayerLevel(Index) + 1) ^ 2) + 17 * (GetPlayerLevel(Index) + 1) - 12)
    Exit Function
        
' Error handler
ErrorHandler:
    HandleError "GetPlayerNextLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Player(Index).exp
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal exp As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    
    Player(Index).exp = exp
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "SetPlayerExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerSkill(ByVal Index As Long, ByVal SkillNum As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerSkill = Player(Index).Skills(SkillNum).Level
    Exit Function
        
' Error handler
ErrorHandler:
    HandleError "GetPlayerSkill", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerSkill(ByVal Index As Long, ByVal Level As Byte, ByVal SkillNum As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    
    Player(Index).Skills(SkillNum).Level = Level
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "SetPlayerSkill", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerNextSkillLevel(ByVal Index As Long, ByVal SkillNum As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    GetPlayerNextSkillLevel = (50 / 3) * ((GetPlayerSkill(Index, SkillNum) + 1) ^ 3 - (6 * (GetPlayerSkill(Index, SkillNum) + 1) ^ 2) + 17 * (GetPlayerSkill(Index, SkillNum) + 1) - 12)
    Exit Function
        
' Error handler
ErrorHandler:
    HandleError "GetPlayerNextSkillLevel", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Function GetPlayerSkillExp(ByVal Index As Long, ByVal SkillNum As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSkillExp = Player(Index).Skills(SkillNum).exp
    Exit Function
        
' Error handler
ErrorHandler:
    HandleError "GetPlayerSkillExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerSkillExp(ByVal Index As Long, ByVal exp As Long, ByVal SkillNum As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Skills(SkillNum).exp = exp
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "SetPlayerSkillExp", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Byte
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Player(Index).Access
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Access = Access
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerAccess", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerPK(ByVal Index As Long) As Byte
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Player(Index).PK
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).PK = PK
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerPK", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Player(Index).Vital(Vital)
    Exit Function
    
ErrorHandler:
    HandleError "GetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Player(Index).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerMaxVital = Player(Index).MaxVital(Vital)
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerMaxVital", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Function GetPlayerStat(ByVal Index As Long, Stat As Stats) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerStat = Player(Index).Stat(Stat)
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerStat(ByVal Index As Long, Stat As Stats, ByVal Value As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If Value <= 0 Then Value = 1
    If Value > MAX_STAT Then Value = MAX_STAT
    Player(Index).Stat(Stat) = Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerStat", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerPOINTS = Player(Index).Points
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal Points As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Points = Points
    Exit Sub
     
' Error handler
ErrorHandler:
    HandleError "SetPlayerPOINTS", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Player(Index).Map
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Map = MapNum
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerMap", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Player(Index).X
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If X < 0 Or X > Map.MaxX Then Exit Sub
    Player(Index).X = X
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerX", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Player(Index).Y
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If Y < 0 Or Y > Map.MaxY Then Exit Sub
    Player(Index).Y = Y
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerY", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Player(Index).Dir
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If Dir < 0 Then Exit Sub
    Player(Index).Dir = Dir
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerDir", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    If InvSlot = 0 Then Exit Function
    GetPlayerInvItemNum = PlayerInv(InvSlot).num
    Exit Function
     
' Error handler
ErrorHandler:
    HandleError "GetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Byte, ByVal ItemNum As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Sub
    PlayerInv(InvSlot).num = ItemNum
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerInvItemNum", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = PlayerInv(InvSlot).Value
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Byte, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Sub
    PlayerInv(InvSlot).Value = ItemValue
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerInvItemValue", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerInvItemBind(ByVal Index As Long, ByVal InvSlot As Byte) As Integer
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemBind = Player(Index).Inv(InvSlot).Bind
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerInvItemBind", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerInvItemBind(ByVal Index As Long, ByVal InvSlot As Byte, ByVal BindType As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Sub
    Player(Index).Inv(InvSlot).Bind = BindType
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerInvItemBind", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Byte) As Byte
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Player(Index).Equipment(EquipmentSlot).num
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal ItemNum As Long, ByVal EquipmentSlot As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot).num = ItemNum
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerEquipment", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerEquipmentDur(ByVal Index As Long, ByVal EquipmentSlot As Byte) As Integer
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerEquipmentDur = Player(Index).Equipment(EquipmentSlot).Durability
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerEquipmentDur", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerEquipmentDur(ByVal Index As Long, ByVal DurValue As Integer, ByVal EquipmentSlot As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot).Durability = DurValue
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerEquipmentDur", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerEquipmentBind(ByVal Index As Long, ByVal EquipmentSlot As Byte) As Byte
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentBind = Player(Index).Equipment(EquipmentSlot).Bind
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerEquipmentBind", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerEquipmentBind(ByVal Index As Long, ByVal BindType As Byte, ByVal EquipmentSlot As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Equipment(EquipmentSlot).Bind = BindType
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerEquipmentBind", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerGender(ByVal Index As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerGender = Player(Index).Gender
    Exit Function
   
' Error handler
ErrorHandler:
    HandleError "GetPlayerGender", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerGender(ByVal Index As Long, GenderNum As Byte)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Player(Index).Gender = GenderNum
    Exit Sub
   
' Error handler
ErrorHandler:
    HandleError "SetPlayerGender", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Byte) As Integer
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemDur = Player(Index).Inv(InvSlot).Durability
    Exit Function
   
' Error handler
ErrorHandler:
    HandleError "GetPlayerInvItemDur", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Byte, ByVal ItemDur As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Sub
    
    Player(Index).Inv(InvSlot).Durability = ItemDur
    Exit Sub
   
' Error handler
ErrorHandler:
    HandleError "SetPlayerInvItemDur", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerHDSerial() As String
    Dim Serial As Long
    Dim VName  As String
    Dim FSName As String

    VName = String$(255, vbNullChar)
    FSName = String$(255, vbNullChar)
    GetVolumeInformation "C:\", VName, 255, Serial, 0, 0, FSName, 255
    VName = Left$(VName, InStr(1, VName, vbNullChar) - 1)
    FSName = Left$(FSName, InStr(1, FSName, vbNullChar) - 1)
    GetPlayerHDSerial = Trim$(str$(Serial))
    Exit Function
   
' Error handler
ErrorHandler:
    HandleError "GetPlayerHDSerial", "modDatabase", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function
