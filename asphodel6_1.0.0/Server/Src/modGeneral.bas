Attribute VB_Name = "modGeneral"
Option Explicit

' ------------------------------------------
' --              Asphodel 6              --
' ------------------------------------------

Private Sub Main()
Dim i As Long
Dim F As Long
Dim time1 As Currency
Dim time2 As Currency

    time1 = GetTickCountNew
    
    MAX_LEVELS = Val(GetVar(App.Path & "\data\config.ini", "MAX_VALUES", "Levels"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "Logging")) <> 1 Then PutVar App.Path & "\data\config.ini", "SETUP", "Logging", CStr(0)
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "Logging")) = 1 Then frmServer.mnuConsoleLog.Checked = True Else frmServer.mnuConsoleLog.Checked = False
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "PVP_Level")) < 1 Or Val(GetVar(App.Path & "\data\config.ini", "SETUP", "PVP_Level")) > MAX_LEVELS Then PutVar App.Path & "\data\config.ini", "SETUP", "PVP_Level", CStr(10)
    frmServer.scrlLevelLimit.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "PVP_Level"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "PVP_LevelOn")) <> 1 Then PutVar App.Path & "\data\config.ini", "SETUP", "PVP_LevelOn", CStr(0)
    frmServer.chkPVPLevel.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "PVP_LevelOn"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "Staff_Safe")) <> 1 Then PutVar App.Path & "\data\config.ini", "SETUP", "Staff_Safe", CStr(0)
    frmServer.chkAdminSafety.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "Staff_Safe"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "SellBack")) < 0 Or Val(GetVar(App.Path & "\data\config.ini", "SETUP", "SellBack")) > 100 Then PutVar App.Path & "\data\config.ini", "SETUP", "SellBack", CStr(50)
    frmServer.scrlSellBack.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "SellBack"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "MapChat")) < 0 Or Val(GetVar(App.Path & "\data\config.ini", "SETUP", "MapChat")) > 1 Then PutVar App.Path & "\data\config.ini", "SETUP", "MapChat", CStr(1)
    frmServer.chkMapChat.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "MapChat"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "GlobalChat")) < 0 Or Val(GetVar(App.Path & "\data\config.ini", "SETUP", "GlobalChat")) > 1 Then PutVar App.Path & "\data\config.ini", "SETUP", "GlobalChat", CStr(1)
    frmServer.chkGlobalChat.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "GlobalChat"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "PrivateChat")) < 0 Or Val(GetVar(App.Path & "\data\config.ini", "SETUP", "PrivateChat")) > 1 Then PutVar App.Path & "\data\config.ini", "SETUP", "PrivateChat", CStr(1)
    frmServer.chkPrivateChat.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "PrivateChat"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "EmoteChat")) < 0 Or Val(GetVar(App.Path & "\data\config.ini", "SETUP", "EmoteChat")) > 1 Then PutVar App.Path & "\data\config.ini", "SETUP", "EmoteChat", CStr(1)
    frmServer.chkEmoteChat.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "EmoteChat"))
    
    If Val(GetVar(App.Path & "\data\config.ini", "SETUP", "StaffOnly")) < 0 Or Val(GetVar(App.Path & "\data\config.ini", "SETUP", "StaffOnly")) > 1 Then PutVar App.Path & "\data\config.ini", "SETUP", "StaffOnly", CStr(0)
    frmServer.chkStaffOnly.Value = Val(GetVar(App.Path & "\data\config.ini", "SETUP", "StaffOnly"))
    
    frmServer.txtNews.Text = GetVar(App.Path & "\data\news.ini", "CONTENT", "News")
    GAME_NEWS = frmServer.txtNews.Text
    
    frmServer.Show
    
    ' Initialize the random-number generator
    Randomize
    
    ' Check if the directory is there, if its not make it
    If LCase$(Dir$(App.Path & "\Data\items", vbDirectory)) <> "items" Then
        Call MkDir(App.Path & "\Data\items")
    End If
    
    If LCase$(Dir$(App.Path & "\Data\maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\Data\maps")
    End If
    
    If LCase$(Dir$(App.Path & "\Data\npcs", vbDirectory)) <> "npcs" Then
        Call MkDir(App.Path & "\Data\npcs")
    End If
    
    If LCase$(Dir$(App.Path & "\Data\shops", vbDirectory)) <> "shops" Then
        Call MkDir(App.Path & "\Data\shops")
    End If
    
    If LCase$(Dir$(App.Path & "\Data\spells", vbDirectory)) <> "spells" Then
        Call MkDir(App.Path & "\Data\spells")
    End If
    
    If LCase$(Dir$(App.Path & "\accounts", vbDirectory)) <> "accounts" Then
        Call MkDir(App.Path & "\accounts")
    End If
    
    If LCase$(Dir$(App.Path & "\Data\signs", vbDirectory)) <> "signs" Then
        Call MkDir(App.Path & "\Data\signs")
    End If
    
    If LCase$(Dir$(App.Path & "\Data\guilds", vbDirectory)) <> "guilds" Then
        Call MkDir(App.Path & "\Data\guilds")
    End If
    
    If LCase$(Dir$(App.Path & "\Data\anims", vbDirectory)) <> "anims" Then
        Call MkDir(App.Path & "\Data\anims")
    End If
    
    Call SetStatus("Loading server options...")
    
    ' used for parsing packets
    SEP_CHAR = vbNullChar ' ChrW$(0)
    END_CHAR = ChrW$(237)
    
    ' set quote character
    vbQuote = ChrW$(34) ' "
    
    ' default server logging on is true
    ServerLog = True
    
    ' set MOTD
    MOTD = GetVar(App.Path & "\data\motd.ini", "MOTD", "Msg")
    
    Load_GameOptions
    Load_BanTable
    
    ' Get the listening socket ready to go
    frmServer.Socket(0).RemoteHost = frmServer.Socket(0).LocalIP
    frmServer.Socket(0).LocalPort = GAME_PORT
    
    ' Init all the player sockets
    Call SetStatus("Initializing player array...")
    
    For i = 1 To MAX_PLAYERS
        Call ClearPlayer(i)
        Load frmServer.Socket(i)
        DoEvents
    Next
    
    ' Serves as a constructor
    Call ClearGameData
    
    Call LoadGameData
    
    If Guild_Creation_Item > 0 Then
        If Item(Guild_Creation_Item).Type <> ItemType.Currency_ Then Guild_Creation_Item = 0: Guild_Creation_Cost = 0
    End If
    
    PutVar App.Path & "\data\config.ini", "GUILD_CONFIG", "ItemNum", CStr(Guild_Creation_Item)
    PutVar App.Path & "\data\config.ini", "GUILD_CONFIG", "Cost", CStr(Guild_Creation_Cost)
    
    Call SetStatus("Spawning map items...")
    Call SpawnAllMapsItems
    
    Call SetStatus("Spawning map npcs...")
    Call SpawnAllMapNpcs
    
    Call SetStatus("Creating map cache...")
    Call CreateFullMapCache
    
    ' Retrieve actual IP address
    ACTUAL_IP = GetActualServerIP
    
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("accounts\charlist.txt") Then
        F = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Output As #F
        Close #F
    End If
    
    ' Start listening
    frmServer.Socket(0).Listen
    
    Call SetStatus("Loading System Tray...")
    Call LoadSystemTray
    
    UpdateCaption
    
    time2 = GetTickCountNew
    
    ' tell the server holder what chats are enabled / disabled
    SetStatus Replace$(Replace$("Map chat: " & frmServer.chkMapChat.Value & "; Global chat: " & frmServer.chkGlobalChat.Value & vbNewLine & _
                                "Private chat: " & frmServer.chkPrivateChat.Value & "; Emote chat: " & frmServer.chkEmoteChat.Value, "1", "ON", , , vbTextCompare), "0", "OFF", , , vbTextCompare)
    
    Call SetStatus("Initialization complete. Server loaded in " & time2 - time1 & "ms.")
    
    frmServer.SSTab1.Enabled = True
    frmServer.mnuServer.Enabled = True
    frmServer.mnuDatabase.Enabled = True
    frmServer.txtChat.Enabled = True
    frmServer.txtChat.SetFocus
    
    Call SetStatus("TIP: Click on the IP at the top to copy it!")
    
    ' Starts the server loop
    ServerLoop
    
End Sub

Public Sub DestroyServer()
Dim i As Long
    
    Call SetStatus("Destroying System Tray...")
    Call DestroySystemTray
    
    Call SetStatus("Saving players online...")
    Call SaveAllPlayersOnline
    
    Call ClearGameData
    
    Call SetStatus("Unloading sockets...")
    
    For i = 1 To MAX_PLAYERS
        Unload frmServer.Socket(i)
    Next
    
    End
    
End Sub

Public Sub CreateFullMapCache()
Dim i As Long

    For i = 1 To MAX_MAPS
        MapCache_Create i
        DoEvents
    Next

End Sub

Public Sub SetStatus(ByVal Status As String)
    Call TextAdd(frmServer.txtText, Status)
    DoEvents
End Sub

Public Sub ClearGameData()
    Call SetStatus("Clearing temp tile fields...")
    Call ClearTempTile
    Call SetStatus("Clearing maps...")
    Call ClearMaps
    Call SetStatus("Clearing map items...")
    Call ClearMapItems
    Call SetStatus("Clearing map npcs...")
    Call ClearMapNpcs
    Call SetStatus("Clearing npcs...")
    Call ClearNpcs
    Call SetStatus("Clearing items...")
    Call ClearItems
    Call SetStatus("Clearing shops...")
    Call ClearShops
    Call SetStatus("Clearing spells...")
    Call ClearSpells
    Call SetStatus("Clearing signs...")
    Call ClearSigns
    Call SetStatus("Clearing guilds...")
    Call ClearGuilds
    Call SetStatus("Clearing animations...")
    Call ClearAnims
End Sub

Public Sub LoadGameData()
    Call SetStatus("Loading classes...")
    Call LoadClasses
    Call SetStatus("Loading maps...")
    Call LoadMaps
    Call SetStatus("Loading items...")
    Call LoadItems
    Call SetStatus("Loading npcs...")
    Call LoadNpcs
    Call SetStatus("Loading shops...")
    Call LoadShops
    Call SetStatus("Loading spells...")
    Call LoadSpells
    Call SetStatus("Loading signs...")
    Call LoadSigns
    Call SetStatus("Loading guilds...")
    Call LoadGuilds
    Call SetStatus("Loading animations...")
    Call LoadAnims
End Sub

Public Sub DestroySystemTray()
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = GAME_NAME & vbNullChar
    Call Shell_NotifyIcon(NIM_DELETE, nid) ' Delete from the sys tray
End Sub

Public Sub LoadSystemTray()
    nid.cbSize = Len(nid)
    nid.hWnd = frmServer.hWnd
    nid.uId = vbNull
    nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    nid.uCallBackMessage = WM_MOUSEMOVE
    nid.hIcon = frmServer.Icon
    nid.szTip = GAME_NAME & vbNullChar   'You can add your game name or something.
    Call Shell_NotifyIcon(NIM_ADD, nid) 'Add to the sys tray
End Sub

' Used for checking validity of names
Function isNameLegal(ByVal sInput As Integer) As Boolean
    isNameLegal = ((sInput >= 65 And sInput <= 90) Or (sInput >= 97 And sInput <= 122) Or (sInput = 95) Or (sInput = 32) Or (sInput >= 48 And sInput <= 57))
End Function

Function Random(Lowerbound As Integer, Upperbound As Integer) As Integer
    Random = Int((Upperbound - Lowerbound + 1) * Rnd) + Lowerbound
End Function

Public Sub UsersOnline_Start()
Dim i As Integer

    For i = 1 To MAX_PLAYERS
        frmServer.lstPlayers.AddItem i & ") None"
    Next
    
End Sub

Public Sub UpdatePlayerTable(ByVal Index As Long)
    If Player(Index).Char(TempPlayer(Index).CharNum).Muted Then
        frmServer.lstPlayers.List(Index - 1) = Index & ") Char: " & GetPlayerName(Index) & " | Account: " & GetPlayerLogin(Index) & " | " & GetPlayerIP(Index) & " | Access " & GetPlayerAccess(Index) & " | Muted |"
    Else
        frmServer.lstPlayers.List(Index - 1) = Index & ") Char: " & GetPlayerName(Index) & " | Account: " & GetPlayerLogin(Index) & " | " & GetPlayerIP(Index) & " | Access " & GetPlayerAccess(Index) & " | Not Muted |"
    End If
End Sub

Public Function GetTickCountNew() As Currency
    GetSysTimeMS GetTickCountNew
End Function

''''''''''''''''''''''''''''''''''''''''''''
' Function written from scratch by: GIAKEN '
''''''''''''''''''''''''''''''''''''''''''''
Public Function GetActualServerIP() As String
Dim RetryValue As Long

    Load frmServer.INet(1)
    
Retry:
    
    'after 2 seconds time the request out
    frmServer.INet(RetryValue).RequestTimeout = 2
    
    Select Case IP_Source
        Case "org"
            Call SetStatus("Retrieving actual IP from whatismyip.org...")
            
            On Error Resume Next
            GetActualServerIP = frmServer.INet(RetryValue).OpenURL("http://whatismyip.org", icString)
            On Error GoTo ErrHandler
            
        Case "com"
            Call SetStatus("Retrieving actual IP from whatismyip.com...")
            
            On Error Resume Next
            GetActualServerIP = frmServer.INet(RetryValue).OpenURL("http://whatismyip.com/automation/n09230945.asp", icString)
            On Error GoTo ErrHandler
            
    End Select
    
    'if the IP is blank, then handle it as an error
    If LenB(GetActualServerIP) = 0 Then GoTo ErrHandler
    
    Call SetStatus("Successfully retrieved actual IP.")
    Exit Function

ErrHandler:

    'clear the RTE error just in case so we can properly continue
    Err.Clear
    
    Call SetStatus("Failed to retrieve actual IP address: timed out.")
    
    frmServer.INet(RetryValue).Cancel
    RetryValue = RetryValue + 1
    
    If RetryValue <= 1 Then
        ' since we failed, we'll retry with another method
        If IP_Source = "org" Then
            IP_Source = "com"
        Else
            IP_Source = "org"
        End If
        GoTo Retry
    End If
    
    'couldn't get the actual IP, so just use the local one...
    GetActualServerIP = frmServer.Socket(0).LocalIP
    
End Function

Public Function IsIP(ByVal IPAddress As String) As Boolean
Dim s() As String
Dim i As Long
    
    'If there are no periods, I have no idea what we have...
    If InStr(1, IPAddress, ".") = 0 Then Exit Function
    
    'Split up the string by the periods
    s = Split(IPAddress, ".")
    
    'Confirm we have ubound = 3, since xxx.xxx.xxx.xxx has 4 elements and we start at index 0
    If UBound(s) <> 3 Then Exit Function
    
    'Check that the values are numeric and in a valid range
    For i = 0 To 3
        If Not IsNumeric(s(i)) Then Exit Function
        If s(i) < 0 Then Exit Function
        If s(i) > 255 Then Exit Function
    Next
    
    'Looks like we were passed a valid IP!
    IsIP = True
    
End Function
