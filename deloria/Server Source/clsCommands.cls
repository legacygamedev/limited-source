VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
    
    File = App.Path & "\" & File
    
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    File = App.Path & "\" & File
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Function IsConnected(ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    Else
        IsConnected = False
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsConnected", Err.Number, Err.Description
End Function

Function IsPlaying(ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
    If IsConnected(index) And Player(index).InGame = True Then
        IsPlaying = True
    Else
        IsPlaying = False
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsPlaying", Err.Number, Err.Description
End Function

Function IsLoggedIn(ByVal index As Long) As Boolean
On Error GoTo ErrorHandler
    If IsConnected(index) And Trim(Player(index).Login) <> "" Then
        IsLoggedIn = True
    Else
        IsLoggedIn = False
    End If
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsLoggedIn", Err.Number, Err.Description
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
On Error GoTo ErrorHandler
Dim i As Long

    IsMultiAccounts = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And LCase(Trim(Player(i).Login)) = LCase(Trim(Login)) Then
            IsMultiAccounts = True
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsMultiAccounts", Err.Number, Err.Description
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
On Error GoTo ErrorHandler
Dim i As Long
Dim n As Long

    n = 0
    IsMultiIPOnline = False
    For i = 1 To MAX_PLAYERS
        If IsConnected(i) And Trim(GetPlayerIP(i)) = Trim(IP) Then
            n = n + 1
            
            If (n > 1) Then
                IsMultiIPOnline = True
                Exit Function
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsMultiIPOnline", Err.Number, Err.Description
End Function

Function IsBanned(ByVal IP As String) As Boolean
On Error GoTo ErrorHandler

Dim FileName As String, fIP As String, fName As String
Dim f As Long

    IsBanned = False
    
    FileName = App.Path & "\banlist.txt"
    
    ' Check if file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If
    
    f = FreeFile
    Open FileName For Input As #f
    
    Do While Not EOF(f)
        Input #f, fIP
        Input #f, fName
    
        ' Is banned?
        If Trim(LCase(fIP)) = Trim(LCase(Mid(IP, 1, Len(fIP)))) Then
            IsBanned = True
            Close #f
            Exit Function
        End If
    Loop
    
    Close #f
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modServerTCP.bas", "IsBanned", Err.Number, Err.Description
End Function

Sub SendDataTo(ByVal index As Long, ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long, n As Long, startc As Long

    If IsConnected(index) Then
        frmServer.Socket(index).SendData Data
        DoEvents
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataTo", Err.Number, Err.Description
End Sub

Sub SendDataToAll(ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataToAll", Err.Number, Err.Description
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And i <> index Then
            Call SendDataTo(i, Data)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataToAllBut", Err.Number, Err.Description
End Sub

Sub SendDataToMap(ByVal MapNum As Long, ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataToMap", Err.Number, Err.Description
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Long, ByVal Data As String)
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum And i <> index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendDataToMapBut", Err.Number, Err.Description
End Sub

Sub GlobalMsg(ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "GLOBALMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToAll(Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "GlobalMsg", Err.Number, Err.Description
End Sub

Sub AdminMsg(ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "ADMINMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerAccess(i) > 0 Then
            Call SendDataTo(i, Packet)
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "AdminMsg", Err.Number, Err.Description
End Sub

Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "PlayerMsg", Err.Number, Err.Description
End Sub

Sub MapMsg(ByVal MapNum As Long, ByVal Msg As String, ByVal Color As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim Text As String

    Packet = "MAPMSG" & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR
    
    Call SendDataToMap(MapNum, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "MapMsg", Err.Number, Err.Description
End Sub

Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "ALERTMSG" & SEP_CHAR & Msg & SEP_CHAR & END_CHAR
    
    Call SendDataTo(index, Packet)
    Call CloseSocket(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "AlertMsg", Err.Number, Err.Description
End Sub

Sub HackingAttempt(ByVal index As Long, ByVal Reason As String)
On Error GoTo ErrorHandler
    If index > 0 Then
        If IsPlaying(index) Then
            Call GlobalMsg(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has been booted for (" & Reason & ")", White)
        End If
    
        Call AlertMsg(index, "You have lost your connection with " & GAME_NAME & ".")
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "HackingAttempt", Err.Number, Err.Description
End Sub

Function GetPlayerLogin(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerLogin = Trim(Player(index).Login)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerLogin", Err.Number, Err.Description
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
On Error GoTo ErrorHandler
    Player(index).Login = Login
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerLogin", Err.Number, Err.Description
End Sub

Function GetPlayerPassword(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerPassword = Trim(Player(index).Password)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerPassword", Err.Number, Err.Description
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
On Error GoTo ErrorHandler
    Player(index).Password = Password
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerPassword", Err.Number, Err.Description
End Sub

Function GetPlayerName(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerName = Trim(Player(index).Char(Player(index).CharNum).Name)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerName", Err.Number, Err.Description
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Name = Name
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerName", Err.Number, Err.Description
End Sub

Function GetPlayerGuild(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerGuild = Trim(Player(index).Char(Player(index).CharNum).Guild)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerGuild", Err.Number, Err.Description
End Function

Sub SetPlayerGuild(ByVal index As Long, ByVal Guild As String)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Guild = Guild
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerGuild", Err.Number, Err.Description
End Sub

Function GetPlayerGuildAccess(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerGuildAccess = Player(index).Char(Player(index).CharNum).Guildaccess
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerGuildAccess", Err.Number, Err.Description
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Guildaccess As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Guildaccess = Guildaccess
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerGuildAccess", Err.Number, Err.Description
End Sub

Function GetPlayerClass(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerClass = Player(index).Char(Player(index).CharNum).Class
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerClass", Err.Number, Err.Description
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Class = ClassNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerClass", Err.Number, Err.Description
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerSprite = Player(index).Char(Player(index).CharNum).Sprite
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSprite", Err.Number, Err.Description
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Sprite = Sprite
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSprite", Err.Number, Err.Description
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerLevel = Player(index).Char(Player(index).CharNum).Level
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerLevel", Err.Number, Err.Description
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Level = Level
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerLevel", Err.Number, Err.Description
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerNextLevel = Experience(GetPlayerLevel(index))
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerNextLevel", Err.Number, Err.Description
End Function

Function GetPlayerExp(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerExp = Player(index).Char(Player(index).CharNum).Exp
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerExp", Err.Number, Err.Description
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Exp = Exp
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerExp", Err.Number, Err.Description
End Sub

Function GetPlayerAccess(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerAccess = Player(index).Char(Player(index).CharNum).Access
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerAccess", Err.Number, Err.Description
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Access = Access
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerAccess", Err.Number, Err.Description
End Sub

Function GetPlayerPK(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerPK = Player(index).Char(Player(index).CharNum).PK
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerPK", Err.Number, Err.Description
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).PK = PK
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerPK", Err.Number, Err.Description
End Sub

Function GetPlayerHP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerHP = Player(index).Char(Player(index).CharNum).HP
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerHP", Err.Number, Err.Description
End Function

Sub SetPlayerHP(ByVal index As Long, ByVal HP As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).HP = HP
    
    If GetPlayerHP(index) > GetPlayerMaxHP(index) Then
        Player(index).Char(Player(index).CharNum).HP = GetPlayerMaxHP(index)
    End If
    If GetPlayerHP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).HP = 0
    End If
    Call SendStats(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerHP", Err.Number, Err.Description
End Sub

Function GetPlayerMP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerMP = Player(index).Char(Player(index).CharNum).MP
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMP", Err.Number, Err.Description
End Function

Sub SetPlayerMP(ByVal index As Long, ByVal MP As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).MP = MP

    If GetPlayerMP(index) > GetPlayerMaxMP(index) Then
        Player(index).Char(Player(index).CharNum).MP = GetPlayerMaxMP(index)
    End If
    If GetPlayerMP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).MP = 0
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerMP", Err.Number, Err.Description
End Sub

Function GetPlayerSP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerSP = Player(index).Char(Player(index).CharNum).SP
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSP", Err.Number, Err.Description
End Function

Sub SetPlayerSP(ByVal index As Long, ByVal SP As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).SP = SP

    If GetPlayerSP(index) > GetPlayerMaxSP(index) Then
        Player(index).Char(Player(index).CharNum).SP = GetPlayerMaxSP(index)
    End If
    If GetPlayerSP(index) < 0 Then
        Player(index).Char(Player(index).CharNum).SP = 0
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSP", Err.Number, Err.Description
End Sub

Function GetPlayerMaxHP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim CharNum As Long
Dim i As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddHP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddHP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddHP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddHP
    End If

    CharNum = Player(index).CharNum
    GetPlayerMaxHP = GetPlayerVIT(index) + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMaxHP", Err.Number, Err.Description
End Function

Function GetPlayerMaxMP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim CharNum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddMP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddMP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddMP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddMP
    End If

    CharNum = Player(index).CharNum
    GetPlayerMaxMP = GetPlayerMAGI(index) + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMaxMP", Err.Number, Err.Description
End Function

Function GetPlayerMaxSP(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim CharNum As Long
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddSP
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddSP
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddSP
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddSP
    End If

    CharNum = Player(index).CharNum
    GetPlayerMaxSP = GetPlayerSPEED(index) + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMaxSP", Err.Number, Err.Description
End Function

Function GetClassName(ByVal ClassNum As Long) As String
On Error GoTo ErrorHandler
    GetClassName = Trim(Class(ClassNum).Name)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassName", Err.Number, Err.Description
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassMaxHP = (1 + Int(Class(ClassNum).STR / 2) + Class(ClassNum).STR) * 2
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassMaxHP", Err.Number, Err.Description
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassMaxMP = (1 + Int(Class(ClassNum).MAGI / 2) + Class(ClassNum).MAGI) * 2
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassMaxMP", Err.Number, Err.Description
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassMaxSP = (1 + Int(Class(ClassNum).SPEED / 2) + Class(ClassNum).SPEED) * 2
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassMaxSP", Err.Number, Err.Description
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassSTR = Class(ClassNum).STR
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassSTR", Err.Number, Err.Description
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassDEF = Class(ClassNum).DEF
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassDEF", Err.Number, Err.Description
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassSPEED = Class(ClassNum).SPEED
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassSPEED", Err.Number, Err.Description
End Function

Function GetClassVIT(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassVIT = Class(ClassNum).VIT
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassVIT", Err.Number, Err.Description
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
On Error GoTo ErrorHandler
    GetClassMAGI = Class(ClassNum).MAGI
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetClassMAGI", Err.Number, Err.Description
End Function

Function GetPlayerBaseSTR(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseSTR = Player(index).Char(Player(index).CharNum).STR
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseSTR", Err.Number, Err.Description
End Function

Function GetPlayerBaseDEF(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseDEF = Player(index).Char(Player(index).CharNum).DEF
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseDEF", Err.Number, Err.Description
End Function

Function GetPlayerBaseMAGI(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseMAGI = Player(index).Char(Player(index).CharNum).MAGI
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseMAGI", Err.Number, Err.Description
End Function

Function GetPlayerBaseSPEED(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseSPEED = Player(index).Char(Player(index).CharNum).SPEED
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseSPEED", Err.Number, Err.Description
End Function

Function GetPlayerBaseVIT(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBaseVIT = Player(index).Char(Player(index).CharNum).VIT
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBaseVIT", Err.Number, Err.Description
End Function

Function GetPlayerSTR(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddStr
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddStr
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddStr
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddStr
    End If
    GetPlayerSTR = Player(index).Char(Player(index).CharNum).STR + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSTR", Err.Number, Err.Description
End Function

Sub SetPlayerSTR(ByVal index As Long, ByVal STR As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).STR = STR
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSTR", Err.Number, Err.Description
End Sub

Function GetPlayerDEF(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddDef
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddDef
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddDef
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddDef
    End If
    GetPlayerDEF = Player(index).Char(Player(index).CharNum).DEF + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerDEF", Err.Number, Err.Description
End Function

Sub SetPlayerDEF(ByVal index As Long, ByVal DEF As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).DEF = DEF
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerDEF", Err.Number, Err.Description
End Sub

Function GetPlayerVIT(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).Data3
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).Data3
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).Data3
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).Data3
    End If
    GetPlayerVIT = Player(index).Char(Player(index).CharNum).VIT + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerVIT", Err.Number, Err.Description
End Function

Sub SetPlayerVIT(ByVal index As Long, ByVal VIT As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).VIT = VIT
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerVIT", Err.Number, Err.Description
End Sub

Function GetPlayerSPEED(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddSpeed
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddSpeed
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddSpeed
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddSpeed
    End If
    GetPlayerSPEED = Player(index).Char(Player(index).CharNum).SPEED + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSPEED", Err.Number, Err.Description
End Function

Sub SetPlayerSPEED(ByVal index As Long, ByVal SPEED As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).SPEED = SPEED
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSPEED", Err.Number, Err.Description
End Sub

Function GetPlayerMAGI(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim add As Long
add = 0
    If GetPlayerWeaponSlot(index) > 0 Then
        add = Item(GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index))).AddMagi
    End If
    If GetPlayerArmorSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerArmorSlot(index))).AddMagi
    End If
    If GetPlayerShieldSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerShieldSlot(index))).AddMagi
    End If
    If GetPlayerHelmetSlot(index) > 0 Then
        add = add + Item(GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index))).AddMagi
    End If
    GetPlayerMAGI = Player(index).Char(Player(index).CharNum).MAGI + add
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMAGI", Err.Number, Err.Description
End Function

Sub SetPlayerMAGI(ByVal index As Long, ByVal MAGI As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).MAGI = MAGI
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerMAGI", Err.Number, Err.Description
End Sub

Function GetPlayerPOINTS(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerPOINTS = Player(index).Char(Player(index).CharNum).POINTS
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerPOINTS", Err.Number, Err.Description
End Function

Sub SetPlayerPOINTS(ByVal index As Long, ByVal POINTS As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).POINTS = POINTS
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerPOINTS", Err.Number, Err.Description
End Sub

Function GetPlayerMap(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerMap = Player(index).Char(Player(index).CharNum).Map
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerMap", Err.Number, Err.Description
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Long)
On Error GoTo ErrorHandler
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(index).Char(Player(index).CharNum).Map = MapNum
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerMap", Err.Number, Err.Description
End Sub

Function GetPlayerX(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerX = Player(index).Char(Player(index).CharNum).x
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerX", Err.Number, Err.Description
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).x = x
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerX", Err.Number, Err.Description
End Sub

Function GetPlayerY(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerY = Player(index).Char(Player(index).CharNum).y
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerY", Err.Number, Err.Description
End Function

Sub SetPlayerY(ByVal index As Long, ByVal y As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).y = y
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerY", Err.Number, Err.Description
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerDir = Player(index).Char(Player(index).CharNum).Dir
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerDir", Err.Number, Err.Description
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Dir = Dir
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerDir", Err.Number, Err.Description
End Sub

Function GetPlayerIP(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerIP", Err.Number, Err.Description
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerInvItemNum = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerInvItemNum", Err.Number, Err.Description
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Num = ItemNum
    Call SendInventoryUpdate(index, InvSlot)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerInvItemNum", Err.Number, Err.Description
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerInvItemValue = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerInvItemValue", Err.Number, Err.Description
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Value = ItemValue
    Call SendInventoryUpdate(index, InvSlot)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerInvItemValue", Err.Number, Err.Description
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerInvItemDur = Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerInvItemDur", Err.Number, Err.Description
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Inv(InvSlot).Dur = ItemDur
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerInvItemDur", Err.Number, Err.Description
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerSpell = Player(index).Char(Player(index).CharNum).Spell(SpellSlot)
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerSpell", Err.Number, Err.Description
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Spell(SpellSlot) = SpellNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerSpell", Err.Number, Err.Description
End Sub

Function GetPlayerArmorSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerArmorSlot = Player(index).Char(Player(index).CharNum).ArmorSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerArmorSlot", Err.Number, Err.Description
End Function

Sub SetPlayerArmorSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).ArmorSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerArmorSlot", Err.Number, Err.Description
End Sub

Function GetPlayerWeaponSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerWeaponSlot = Player(index).Char(Player(index).CharNum).WeaponSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerWeaponSlot", Err.Number, Err.Description
End Function

Sub SetPlayerWeaponSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).WeaponSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerWeaponSlot", Err.Number, Err.Description
End Sub

Function GetPlayerHelmetSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerHelmetSlot = Player(index).Char(Player(index).CharNum).HelmetSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerHelmetSlot", Err.Number, Err.Description
End Function

Sub SetPlayerHelmetSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).HelmetSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerHelmetSlot", Err.Number, Err.Description
End Sub

Function GetPlayerShieldSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerShieldSlot = Player(index).Char(Player(index).CharNum).ShieldSlot
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerShieldSlot", Err.Number, Err.Description
End Function

Sub SetPlayerShieldSlot(ByVal index As Long, InvNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).ShieldSlot = InvNum
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerShieldSlot", Err.Number, Err.Description
End Sub

Sub BattleMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, ByVal Who As Long)
On Error GoTo ErrorHandler
    Call SendDataTo(index, "damagedisplay" & SEP_CHAR & Who & SEP_CHAR & Msg & SEP_CHAR & Color & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "BattleMsg", Err.Number, Err.Description
End Sub

Function GetPlayerClassName(ByVal index As Long) As String
On Error GoTo ErrorHandler
    GetPlayerClassName = Class(GetPlayerClass(index)).Name
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerClassName", Err.Number, Err.Description
End Function

Sub PlaySound(ByVal index As Long, ByVal Sound As String)
On Error GoTo ErrorHandler
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "soundattribute" & SEP_CHAR & Sound & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "PlaySound", Err.Number, Err.Description
End Sub

Function Rand(ByVal High As Long, ByVal Low As Long)
On Error GoTo ErrorHandler
Rand = -99999999
Randomize
High = High + 1
Do Until Rand >= Low
    Rand = Int(Rnd * High)
Loop
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "Rand", Err.Number, Err.Description
End Function

Sub SendInventory(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long

    Packet = "PLAYERINV" & SEP_CHAR
    For i = 1 To MAX_INV
        Packet = Packet & GetPlayerInvItemNum(index, i) & SEP_CHAR & GetPlayerInvItemValue(index, i) & SEP_CHAR & GetPlayerInvItemDur(index, i) & SEP_CHAR
    Next i
    Packet = Packet & END_CHAR
    
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendInventory", Err.Number, Err.Description
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Long)
On Error GoTo ErrorHandler
Dim Packet As String
    
    Packet = "PLAYERINVUPDATE" & SEP_CHAR & InvSlot & SEP_CHAR & GetPlayerInvItemNum(index, InvSlot) & SEP_CHAR & GetPlayerInvItemValue(index, InvSlot) & SEP_CHAR & GetPlayerInvItemDur(index, InvSlot) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendInventoryUpdate", Err.Number, Err.Description
End Sub

Sub SendWornEquipment(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
    
    Packet = "PLAYERWORNEQ" & SEP_CHAR & GetPlayerArmorSlot(index) & SEP_CHAR & GetPlayerWeaponSlot(index) & SEP_CHAR & GetPlayerHelmetSlot(index) & SEP_CHAR & GetPlayerShieldSlot(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendWornEquipment", Err.Number, Err.Description
End Sub

Sub SendHP(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERHP" & SEP_CHAR & GetPlayerMaxHP(index) & SEP_CHAR & GetPlayerHP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendHP", Err.Number, Err.Description
End Sub

Sub SendMP(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERMP" & SEP_CHAR & GetPlayerMaxMP(index) & SEP_CHAR & GetPlayerMP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendMP", Err.Number, Err.Description
End Sub

Sub SendSP(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    Packet = "PLAYERSP" & SEP_CHAR & GetPlayerMaxSP(index) & SEP_CHAR & GetPlayerSP(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendSP", Err.Number, Err.Description
End Sub

Sub SendStats(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String
    
    Packet = "PLAYERSTATSPACKET" & SEP_CHAR & GetPlayerSTR(index) & SEP_CHAR & GetPlayerDEF(index) & SEP_CHAR & GetPlayerSPEED(index) & SEP_CHAR & GetPlayerMAGI(index) & SEP_CHAR & GetPlayerNextLevel(index) & SEP_CHAR & GetPlayerExp(index) & SEP_CHAR & GetPlayerLevel(index) & SEP_CHAR & GetPlayerVIT(index) & SEP_CHAR & END_CHAR
    Call SendDataTo(index, Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendStats", Err.Number, Err.Description
End Sub

Sub PlayerPoints(ByVal index As Long)
On Error GoTo ErrorHandler
    Call SendDataTo(index, "playerpoints" & SEP_CHAR & GetPlayerPOINTS(index) & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "PlayerPoints", Err.Number, Err.Description
End Sub

Sub PlayerClass(ByVal index As Long)
On Error GoTo ErrorHandler
    Call SendDataTo(index, "playerclass" & SEP_CHAR & GetPlayerClass(index) & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "PlayerClass", Err.Number, Err.Description
End Sub

Sub SendPlayerData(ByVal index As Long)
On Error GoTo ErrorHandler
Dim Packet As String

    ' Send index's player data to everyone including himself on th emap
    Packet = "PLAYERDATA" & SEP_CHAR
    Packet = Packet & index & SEP_CHAR
    Packet = Packet & GetPlayerName(index) & SEP_CHAR
    Packet = Packet & GetPlayerSprite(index) & SEP_CHAR
    Packet = Packet & GetPlayerMap(index) & SEP_CHAR
    Packet = Packet & GetPlayerX(index) & SEP_CHAR
    Packet = Packet & GetPlayerY(index) & SEP_CHAR
    Packet = Packet & GetPlayerDir(index) & SEP_CHAR
    Packet = Packet & GetPlayerAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerPK(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuild(index) & SEP_CHAR
    Packet = Packet & GetPlayerGuildAccess(index) & SEP_CHAR
    Packet = Packet & GetPlayerClass(index) & SEP_CHAR
    Packet = Packet & Player(index).Char(Player(index).CharNum).Sex & SEP_CHAR
    Packet = Packet & END_CHAR
    Call SendDataToMap(GetPlayerMap(index), Packet)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modServerTCP.bas", "SendPlayerData", Err.Number, Err.Description
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim OldMap As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Check if there was an npc on the map the player is leaving, and if so say goodbye
    'If Trim(Shop(ShopNum).LeaveSay) <> "" Then
        'Call PlayerMsg(Index, Trim(Shop(ShopNum).Name) & " : " & Trim(Shop(ShopNum).LeaveSay) & "", SayColor)
    'End If
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    Call SendLeaveMap(index, OldMap)
    
    Call SetPlayerMap(index, MapNum)
    Call SetPlayerX(index, x)
    Call SetPlayerY(index, y)
                
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
    End If
    
    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES

    Player(index).GettingMap = YES
    Call SendDataToMap(GetPlayerMap(index), "sound" & SEP_CHAR & "warp" & SEP_CHAR & END_CHAR)
    Call SendDataTo(index, "CHECKFORMAP" & SEP_CHAR & MapNum & SEP_CHAR & Map(MapNum).Revision & SEP_CHAR & END_CHAR)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "PlayerWarp", Err.Number, Err.Description
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
On Error GoTo ErrorHandler
Dim x As Long
Dim y As Long
Dim i As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
        
    ' Spawn what we have
    For y = 0 To MAX_MAPY
        For x = 0 To MAX_MAPX
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(x, y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(x, y).Data1).Type = ITEM_TYPE_CURRENCY And Map(MapNum).Tile(x, y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, 1, MapNum, x, y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(x, y).Data1, Map(MapNum).Tile(x, y).Data2, MapNum, x, y)
                End If
            End If
        Next x
    Next y
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnMapItems", Err.Number, Err.Description
End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
On Error GoTo ErrorHandler
Dim i As Long
    
    FindOpenInvSlot = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                If Trim(LCase(Item(GetPlayerInvItemNum(index, i)).Name)) = "gold" Then
                    FindOpenInvSlot = i
                Else
                    If GetPlayerInvItemValue(index, i) >= 100 Then
                        FindOpenInvSlot = 0
                    Else
                        FindOpenInvSlot = i
                    End If
                End If
                Exit Function
            End If
        Next i
    End If
    
    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindOpenInvSlot", Err.Number, Err.Description
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
On Error GoTo ErrorHandler
Dim i As Long

    FindOpenMapItemSlot = 0
    
    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Function
    End If
    
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindOpenMapItemSlot", Err.Number, Err.Description
End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
On Error GoTo ErrorHandler
Dim i As Long

    FindOpenSpellSlot = 0
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindOpenSpellSlot", Err.Number, Err.Description
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
On Error GoTo ErrorHandler
Dim i As Long

    HasSpell = False
    
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "HasSpell", Err.Number, Err.Description
End Function
Function TotalOnlinePlayers() As Long
On Error GoTo ErrorHandler
Dim i As Long
    frmServer.LstPlayers.Clear
    frmServer.LstAccounts.Clear
    TotalOnlinePlayers = 0

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
            frmServer.LstPlayers.AddItem Trim(Player(i).Char(Player(i).CharNum).Name)
            frmServer.LstAccounts.AddItem Trim(Player(i).Login)
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "TotalOnlinePlayers", Err.Number, Err.Description
End Function

Function FindPlayer(ByVal Name As String) As Long
On Error GoTo ErrorHandler
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim(Name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim(Name)))) = UCase(Trim(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "FindPlayer", Err.Number, Err.Description
End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Long) As Long
On Error GoTo ErrorHandler
Dim i As Long
    
    HasItem = 0
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Function
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If
            Exit Function
        End If
    Next i
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modGameLogic.bas", "HasItem", Err.Number, Err.Description
End Function

Sub TakeItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
On Error GoTo ErrorHandler
Dim i As Long, n As Long
Dim TakeItem As Boolean

    TakeItem = False
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    Call SendInventoryUpdate(index, i)
                End If
            Else
                ' Check to see if its any sort of ArmorSlot/WeaponSlot
                Select Case Item(GetPlayerInvItemNum(index, i)).Type
                    Case ITEM_TYPE_WEAPON
                        If GetPlayerWeaponSlot(index) > 0 Then
                            If i = GetPlayerWeaponSlot(index) Then
                                Call SetPlayerWeaponSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerWeaponSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                
                    Case ITEM_TYPE_ARMOR
                        If GetPlayerArmorSlot(index) > 0 Then
                            If i = GetPlayerArmorSlot(index) Then
                                Call SetPlayerArmorSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerArmorSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_HELMET
                        If GetPlayerHelmetSlot(index) > 0 Then
                            If i = GetPlayerHelmetSlot(index) Then
                                Call SetPlayerHelmetSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerHelmetSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                    
                    Case ITEM_TYPE_SHIELD
                        If GetPlayerShieldSlot(index) > 0 Then
                            If i = GetPlayerShieldSlot(index) Then
                                Call SetPlayerShieldSlot(index, 0)
                                Call SendWornEquipment(index)
                                TakeItem = True
                            Else
                                ' Check if the item we are taking isn't already equipped
                                If ItemNum <> GetPlayerInvItemNum(index, GetPlayerShieldSlot(index)) Then
                                    TakeItem = True
                                End If
                            End If
                        Else
                            TakeItem = True
                        End If
                End Select

                
                n = Item(GetPlayerInvItemNum(index, i)).Type
                ' Check if its not an equipable weapon, and if it isn't then take it away
                If (n <> ITEM_TYPE_WEAPON) And (n <> ITEM_TYPE_ARMOR) And (n <> ITEM_TYPE_HELMET) And (n <> ITEM_TYPE_SHIELD) Then
                    TakeItem = True
                End If
            End If
                            
            If TakeItem = True Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Call SetPlayerInvItemDur(index, i, 0)
                
                ' Send the inventory update
                Call SendInventoryUpdate(index, i)
                Exit Sub
            End If
        End If
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "TakeItem", Err.Number, Err.Description
End Sub

Sub GiveItem(ByVal index As Long, ByVal ItemNum As Long, ByVal ItemVal As Long)
On Error GoTo ErrorHandler
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then
        Exit Sub
    End If
    
    i = FindOpenInvSlot(index, ItemNum)
    
    ' Check to see if inventory is full
    If i <> 0 Then
        Call SetPlayerInvItemNum(index, i, ItemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        
        If (Item(ItemNum).Type = ITEM_TYPE_ARMOR) Or (Item(ItemNum).Type = ITEM_TYPE_WEAPON) Or (Item(ItemNum).Type = ITEM_TYPE_HELMET) Or (Item(ItemNum).Type = ITEM_TYPE_SHIELD) Then
            Call SetPlayerInvItemDur(index, i, Item(ItemNum).Data1)
        End If
        
        Call SendInventoryUpdate(index, i)
    Else
        Call BattleMsg(index, "Your inventory is full.", BrightRed, index)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "GiveItem", Err.Number, Err.Description
End Sub

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
On Error GoTo ErrorHandler
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    Call SpawnItemSlot(i, ItemNum, ItemVal, Item(ItemNum).Data1, MapNum, x, y)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnItem", Err.Number, Err.Description
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal ItemDur As Long, ByVal MapNum As Long, ByVal x As Long, ByVal y As Long)
On Error GoTo ErrorHandler
Dim Packet As String
Dim i As Long
    
    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then
        Exit Sub
    End If
    
    i = MapItemSlot
    
    If i <> 0 And ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
        MapItem(MapNum, i).Num = ItemNum
        MapItem(MapNum, i).Value = ItemVal
        
        If ItemNum <> 0 Then
            If (Item(ItemNum).Type >= ITEM_TYPE_WEAPON) And (Item(ItemNum).Type <= ITEM_TYPE_SHIELD) Then
                MapItem(MapNum, i).Dur = ItemDur
            Else
                MapItem(MapNum, i).Dur = 0
            End If
        Else
            MapItem(MapNum, i).Dur = 0
        End If
        
        MapItem(MapNum, i).x = x
        MapItem(MapNum, i).y = y
            
        Packet = "SPAWNITEM" & SEP_CHAR & i & SEP_CHAR & ItemNum & SEP_CHAR & ItemVal & SEP_CHAR & MapItem(MapNum, i).Dur & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & END_CHAR
        Call SendDataToMap(MapNum, Packet)
    End If
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnItemSlot", Err.Number, Err.Description
End Sub

Sub SpawnAllMapsItems()
On Error GoTo ErrorHandler
Dim i As Long
    
    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next i
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "SpawnAllMapsItems", Err.Number, Err.Description
End Sub

Function GetPlayerCharNum(ByVal index As Long) As Long
    GetPlayerCharNum = Player(index).CharNum
End Function

Sub Flash(ByVal index As Long, ByVal flashfile As String)
    Call SendDataTo(index, "flashevent" & SEP_CHAR & flashfile & SEP_CHAR & END_CHAR)
End Sub

Sub Prompt(ByVal index As Long, ByVal question As String, ByVal Value As Long)
    Call SendDataTo(index, "prompt" & SEP_CHAR & question & SEP_CHAR & Value & SEP_CHAR & END_CHAR)
End Sub

Function GetBootMap(ByVal index As Long) As Long
    GetBootMap = Map(GetPlayerMap(index)).BootMap
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Long, ByVal Amount As Long)
On Error GoTo ErrorHandler
Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvNum <= 0 Or InvNum > MAX_INV Then
        Exit Sub
    End If
    
    If (GetPlayerInvItemNum(index, InvNum) > 0) And (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(index))
        
        If i <> 0 Then
            MapItem(GetPlayerMap(index), i).Dur = 0
            
            ' Check to see if its any sort of ArmorSlot/WeaponSlot
            Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum = GetPlayerArmorSlot(index) Then
                        Call SetPlayerArmorSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum = GetPlayerWeaponSlot(index) Then
                        Call SetPlayerWeaponSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                    
                Case ITEM_TYPE_HELMET
                    If InvNum = GetPlayerHelmetSlot(index) Then
                        Call SetPlayerHelmetSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
                                    
                Case ITEM_TYPE_SHIELD
                    If InvNum = GetPlayerShieldSlot(index) Then
                        Call SetPlayerShieldSlot(index, 0)
                        Call SendWornEquipment(index)
                    End If
                    MapItem(GetPlayerMap(index), i).Dur = GetPlayerInvItemDur(index, InvNum)
            End Select
                                
            MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).x = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).y = GetPlayerY(index)
                        
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, InvNum)
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & GetPlayerInvItemValue(index, InvNum) & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).Value = Amount
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops " & Amount & " " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Amount)
                End If
            Else
                ' Its not a currency object so this is easy
                MapItem(GetPlayerMap(index), i).Value = 0
                If Item(GetPlayerInvItemNum(index, InvNum)).Type >= ITEM_TYPE_WEAPON And Item(GetPlayerInvItemNum(index, InvNum)).Type <= ITEM_TYPE_SHIELD Then
                    If Item(GetPlayerInvItemNum(index, InvNum)).Data1 <= -1 Then
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    Else
                        Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                    End If
                Else
                    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " drops a " & Trim(Item(GetPlayerInvItemNum(index, InvNum)).Name) & ".", Yellow)
                End If
                
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
            End If
                                        
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, Amount, MapItem(GetPlayerMap(index), i).Dur, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "To many items already on the ground.", BrightRed)
        End If
    End If
    
    Call SendInventory(index)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modGameLogic.bas", "PlayerMapDropItem", Err.Number, Err.Description
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBankItemNum = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Num
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBankItemNum", Err.Number, Err.Description
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Num = ItemNum
    Call SendBankUpdate(index, BankSlot)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerBankItemNum", Err.Number, Err.Description
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBankItemValue = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBankItemValue", Err.Number, Err.Description
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(index, BankSlot)
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerBankItemValue", Err.Number, Err.Description
End Sub

Function GetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long) As Long
On Error GoTo ErrorHandler
    GetPlayerBankItemDur = Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur
ErrorHandlerExit:
  Exit Function
ErrorHandler:
  ReportError "modTypes.bas", "GetPlayerBankItemDur", Err.Number, Err.Description
End Function

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
On Error GoTo ErrorHandler
    Player(index).Char(Player(index).CharNum).Bank(BankSlot).Dur = ItemDur
ErrorHandlerExit:
  Exit Sub
ErrorHandler:
  ReportError "modTypes.bas", "SetPlayerBankItemDur", Err.Number, Err.Description
End Sub

