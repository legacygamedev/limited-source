Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpFileName As String) As Long

' For Clear functions
Public Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If tName = "" Then Exit Sub
    If LCase$(Dir$(tDir & tName, vbDirectory)) <> LCase$(tName) Then Call MkDir(LCase$(tDir & "\" & tName))
End Sub

' Outputs string to text file
Public Function TimeStamp() As String
    TimeStamp = "[" & Time & "]"
End Function

Public Sub AddLog(ByVal Text As String, ByVal LogFile As String)
    Dim filename As String
    Dim F As Integer

    Call ChkDir(App.path & "\", "logs")
    Call ChkDir(App.path & "\logs\", Month(Now) & "-" & Day(Now) & "-" & Year(Now))
    filename = App.path & "\logs\" & Month(Now) & "-" & Day(Now) & "-" & Year(Now) & "\" & LogFile & ".log"

    If Not FileExist(filename, True) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    
    Open filename For Append As #F
        Print #F, TimeStamp & " - " & Text
    Close #F
End Sub

' Gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' Writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean
    If Not RAW Then
        If Len(Dir$(App.path & filename)) > 0 Then
            FileExist = True
        End If
    Else
        If Len(Dir$(filename)) > 0 Then
            FileExist = True
        End If
    End If
End Function

Public Sub InitOptions()
    Dim filename As String
    
    ' File name used for options
    filename = App.path & "\data\options.ini"
    
    ' Game Name
    If GetVar(filename, "Options", "Name") = "" Then
        Options.Name = "Eclipse Worlds"
        Call PutVar(filename, "Options", "Name", Trim$(Options.Name))
    Else
        Options.Name = GetVar(filename, "Options", "Name")
    End If
    
    ' Website
    If GetVar(filename, "Options", "Website") = "" Then
        Options.Website = "http://www.onlinegamecreation.com/"
        Call PutVar(filename, "Options", "Website", Trim$(Options.Website))
    Else
        Options.Website = GetVar(filename, "Options", "Website")
    End If
    
    ' Port
    If GetVar(filename, "Options", "Port") = "" Then
        Options.Port = "7001"
        Call PutVar(filename, "Options", "Port", Trim$(Options.Port))
    Else
        Options.Port = GetVar(filename, "Options", "Port")
    End If
    
    ' Message of the Day
    If GetVar(filename, "Options", "MOTD") = "" Then
        Options.MOTD = "Welcome to the Legends of Arteix!"
        Call PutVar(filename, "Options", "MOTD", Trim$(Options.MOTD))
    Else
        Options.MOTD = GetVar(filename, "Options", "MOTD")
    End If
    
    ' Staff Message of the Day
    If GetVar(filename, "Options", "SMOTD") = "" Then
        Options.SMOTD = ""
        Call PutVar(filename, "Options", "SMOTD", Trim$(Options.SMOTD))
    Else
        Options.SMOTD = GetVar(filename, "Options", "SMOTD")
    End If

    ' Player Kill level
    If GetVar(filename, "Options", "PKLevel") = "" Then
        Options.PKLevel = "10"
        Call PutVar(filename, "Options", "PKLevel", Trim$(Options.PKLevel))
    Else
        Options.PKLevel = GetVar(filename, "Options", "PKLevel")
    End If
    
    ' Same IP
    If GetVar(filename, "Options", "MultipleIP") = "" Then
        Options.MultipleIP = "1"
        Call PutVar(filename, "Options", "MultipleIP", Trim$(Options.MultipleIP))
    Else
        Options.MultipleIP = GetVar(filename, "Options", "MultipleIP")
    End If
    
    ' Same Serial
    If GetVar(filename, "Options", "MultipleSerial") = "" Then
        Options.MultipleSerial = "1"
        Call PutVar(filename, "Options", "MultipleSerial", Trim$(Options.MultipleSerial))
    Else
        Options.MultipleSerial = GetVar(filename, "Options", "MultipleSerial")
    End If
    
     ' Guild Cost
    If GetVar(filename, "Options", "GuildCost") = "" Then
        Options.GuildCost = "5000"
        Call PutVar(filename, "Options", "GuildCost", Trim$(Options.GuildCost))
    Else
        Options.GuildCost = GetVar(filename, "Options", "GuildCost")
    End If
    
    ' News
    If GetVar(filename, "Options", "News") = "" Then
        Options.News = "Welcome to the Legends of Arteix!"
        Call PutVar(filename, "Options", "News", Trim$(Options.News))
    Else
        Options.News = GetVar(filename, "Options", "News")
    End If
    
    ' Sound
    If GetVar(filename, "Options", "MissSound") = "" Then
        Options.MissSound = "Miss2"
        Call PutVar(filename, "Options", "MissSound", Trim$(Options.MissSound))
    Else
        Options.MissSound = GetVar(filename, "Options", "MissSound")
    End If
    
    If GetVar(filename, "Options", "DodgeSound") = "" Then
        Options.DodgeSound = "Dodge"
        Call PutVar(filename, "Options", "DodgeSound", Trim$(Options.DodgeSound))
    Else
        Options.DodgeSound = GetVar(filename, "Options", "DodgeSound")
    End If
    
    If GetVar(filename, "Options", "DeflectSound") = "" Then
        Options.DeflectSound = "Saint3"
        Call PutVar(filename, "Options", "DeflectSound", Trim$(Options.DeflectSound))
    Else
        Options.DeflectSound = GetVar(filename, "Options", "DeflectSound")
    End If
    
    If GetVar(filename, "Options", "BlockSound") = "" Then
        Options.BlockSound = "Block"
        Call PutVar(filename, "Options", "BlockSound", Trim$(Options.BlockSound))
    Else
        Options.BlockSound = GetVar(filename, "Options", "BlockSound")
    End If
    
    If GetVar(filename, "Options", "CriticalSound") = "" Then
        Options.CriticalSound = "Critical"
        Call PutVar(filename, "Options", "CriticalSound", Trim$(Options.CriticalSound))
    Else
        Options.CriticalSound = GetVar(filename, "Options", "CriticalSound")
    End If
    
    If GetVar(filename, "Options", "ResistSound") = "" Then
        Options.ResistSound = "Saint9"
        Call PutVar(filename, "Options", "ResistSound", Trim$(Options.ResistSound))
    Else
        Options.ResistSound = GetVar(filename, "Options", "ResistSound")
    End If
    
    If GetVar(filename, "Options", "BuySound") = "" Then
        Options.BuySound = "Shop"
        Call PutVar(filename, "Options", "BuySound", Trim$(Options.BuySound))
    Else
        Options.BuySound = GetVar(filename, "Options", "BuySound")
    End If
    
    If GetVar(filename, "Options", "SellSound") = "" Then
        Options.SellSound = "Sell"
        Call PutVar(filename, "Options", "SellSound", Trim$(Options.SellSound))
    Else
        Options.SellSound = GetVar(filename, "Options", "SellSound")
    End If
    
    ' Animations
    If GetVar(filename, "Options", "DeflectAnimation") = "" Then
        Options.DeflectAnimation = 2
        Call PutVar(filename, "Options", "DeflectAnimation", Trim$(Options.DeflectAnimation))
    Else
        Options.DeflectAnimation = GetVar(filename, "Options", "DeflectAnimation")
    End If
    
    If GetVar(filename, "Options", "CriticalAnimation") = "" Then
        Options.CriticalAnimation = 3
        Call PutVar(filename, "Options", "CriticalAnimation", Trim$(Options.CriticalAnimation))
    Else
        Options.CriticalAnimation = GetVar(filename, "Options", "CriticalAnimation")
    End If
    
    If GetVar(filename, "Options", "DodgeAnimation") = "" Then
        Options.DodgeAnimation = 4
        Call PutVar(filename, "Options", "DodgeAnimation", Trim$(Options.DodgeAnimation))
    Else
        Options.DodgeAnimation = GetVar(filename, "Options", "DodgeAnimation")
    End If
    
    If GetVar(filename, "Options", "MaxLevel") = "" Then
        Options.MaxLevel = 99
        Call PutVar(filename, "Options", "MaxLevel", Trim$(Options.MaxLevel))
    Else
        Options.MaxLevel = GetVar(filename, "Options", "MaxLevel")
    End If
    
    If GetVar(filename, "Options", "StatsLevel") = "" Then
        Options.StatsLevel = 5
        Call PutVar(filename, "Options", "StatsLevel", Trim$(Options.StatsLevel))
    Else
        Options.StatsLevel = GetVar(filename, "Options", "StatsLevel")
    End If
    
    If GetVar(filename, "Options", "MaxStat") = "" Then
        Options.MaxStat = 255
        Call PutVar(filename, "Options", "MaxStat", Trim$(Options.MaxStat))
    Else
        Options.MaxStat = GetVar(filename, "Options", "MaxStat")
    End If
    
    If GetVar(filename, "Options", "LevelUpAnimation") = "" Then
        Options.LevelUpAnimation = 7
        Call PutVar(filename, "Options", "LevelUpAnimation", Trim$(Options.LevelUpAnimation))
    Else
        Options.LevelUpAnimation = GetVar(filename, "Options", "LevelUpAnimation")
    End If
    
    If GetVar(filename, "Data Sizes", "Maps") = "" Then
        Call PutVar(filename, "Data Sizes", "Maps", "100")
    End If
    
    If GetVar(filename, "Data Sizes", "NPCs") = "" Then
        Call PutVar(filename, "Data Sizes", "NPCs", "250")
    End If
    
    If GetVar(filename, "Data Sizes", "Animations") = "" Then
        Call PutVar(filename, "Data Sizes", "Animations", "100")
    End If
    
    If GetVar(filename, "Data Sizes", "Events") = "" Then
        Call PutVar(filename, "Data Sizes", "Events", "100")
    End If
    
    If GetVar(filename, "Data Sizes", "Shops") = "" Then
        Call PutVar(filename, "Data Sizes", "Shops", "100")
    End If
    
    If GetVar(filename, "Data Sizes", "Quests") = "" Then
        Call PutVar(filename, "Data Sizes", "Quests", "100")
    End If
    
    If GetVar(filename, "Data Sizes", "Classes") = "" Then
        Call PutVar(filename, "Data Sizes", "Classes", "100")
    End If
    
    If GetVar(filename, "Data Sizes", "Titles") = "" Then
        Call PutVar(filename, "Data Sizes", "Titles", "100")
    End If
    
    If GetVar(filename, "Data Sizes", "Emoticons") = "" Then
        Call PutVar(filename, "Data Sizes", "Emoticons", "50")
    End If
    
    If GetVar(filename, "Data Sizes", "Morals") = "" Then
        Call PutVar(filename, "Data Sizes", "Morals", "25")
    End If
    
    If GetVar(filename, "Data Sizes", "Items") = "" Then
        Call PutVar(filename, "Data Sizes", "Items", "250")
    End If
    
    If GetVar(filename, "Data Sizes", "Bans") = "" Then
        Call PutVar(filename, "Data Sizes", "Bans", "100")
    End If
    
    If GetVar(filename, "Data Sizes", "Resources") = "" Then
        Call PutVar(filename, "Data Sizes", "Resources", "50")
    End If
    
    If GetVar(filename, "Data Sizes", "Spells") = "" Then
        Call PutVar(filename, "Data Sizes", "Spells", "100")
    End If
    
    LoadDataSizes
    
End Sub

Public Sub SaveOptions()
    PutVar App.path & "\data\options.ini", "Options", "Name", Trim$(Options.Name)
    PutVar App.path & "\data\options.ini", "Options", "Port", Trim$(Options.Port)
    PutVar App.path & "\data\options.ini", "Options", "MOTD", Trim$(Options.MOTD)
    PutVar App.path & "\data\options.ini", "Options", "SMOTD", Trim$(Options.SMOTD)
    PutVar App.path & "\data\options.ini", "Options", "Website", Trim$(Options.Website)
    PutVar App.path & "\data\options.ini", "Options", "PKLevel", Trim$(Options.PKLevel)
    PutVar App.path & "\data\options.ini", "Options", "MultipleIP", Trim$(Options.MultipleIP)
    PutVar App.path & "\data\options.ini", "Options", "MultipleSerial", Trim$(Options.MultipleSerial)
    PutVar App.path & "\data\options.ini", "Options", "GuildCost", Trim$(Options.GuildCost)
    PutVar App.path & "\data\options.ini", "Options", "News", Trim$(Options.News)
    PutVar App.path & "\data\options.ini", "Options", "MissSound", Trim$(Options.MissSound)
    PutVar App.path & "\data\options.ini", "Options", "DodgeSound", Trim$(Options.DodgeSound)
    PutVar App.path & "\data\options.ini", "Options", "DeflectSound", Trim$(Options.DeflectSound)
    PutVar App.path & "\data\options.ini", "Options", "BlockSound", Trim$(Options.BlockSound)
    PutVar App.path & "\data\options.ini", "Options", "CriticalSound", Trim$(Options.CriticalSound)
    PutVar App.path & "\data\options.ini", "Options", "ResistSound", Trim$(Options.ResistSound)
    PutVar App.path & "\data\options.ini", "Options", "BuySound", Trim$(Options.BuySound)
    PutVar App.path & "\data\options.ini", "Options", "SellSound", Trim$(Options.SellSound)
    PutVar App.path & "\data\options.ini", "Options", "DeflectAnimation", Trim$(Options.DeflectAnimation)
    PutVar App.path & "\data\options.ini", "Options", "CriticalAnimation", Trim$(Options.CriticalAnimation)
    PutVar App.path & "\data\options.ini", "Options", "DodgeAnimation", Trim$(Options.DodgeAnimation)
    PutVar App.path & "\data\options.ini", "Options", "MaxLevel", Trim$(Options.MaxLevel)
    PutVar App.path & "\data\options.ini", "Options", "StatsLevel", Trim$(Options.StatsLevel)
    PutVar App.path & "\data\options.ini", "Options", "MaxStat", Trim$(Options.MaxStat)
    PutVar App.path & "\data\options.ini", "Options", "LevelUpAnimation", Trim$(Options.LevelUpAnimation)
End Sub

Public Sub LoadOptions()
    Options.Name = GetVar(App.path & "\data\options.ini", "Options", "Name")
    Options.Port = GetVar(App.path & "\data\options.ini", "Options", "Port")
    Options.MOTD = GetVar(App.path & "\data\options.ini", "Options", "MOTD")
    Options.Website = GetVar(App.path & "\data\options.ini", "Options", "Website")
    Options.PKLevel = GetVar(App.path & "\data\options.ini", "Options", "PKLevel")
    Options.MultipleIP = GetVar(App.path & "\data\options.ini", "Options", "MultipleIP")
    Options.MultipleSerial = GetVar(App.path & "\data\options.ini", "Options", "MultipleSerial")
    Options.GuildCost = GetVar(App.path & "\data\options.ini", "Options", "GuildCost")
    Options.News = GetVar(App.path & "\data\options.ini", "Options", "News")
    Options.MissSound = GetVar(App.path & "\data\options.ini", "Options", "MissSound")
    Options.DodgeSound = GetVar(App.path & "\data\options.ini", "Options", "DodgeSound")
    Options.DeflectSound = GetVar(App.path & "\data\options.ini", "Options", "DeflectSound")
    Options.BlockSound = GetVar(App.path & "\data\options.ini", "Options", "BlockSound")
    Options.CriticalSound = GetVar(App.path & "\data\options.ini", "Options", "CriticalSound")
    Options.ResistSound = GetVar(App.path & "\data\options.ini", "Options", "ResistSound")
    Options.BuySound = GetVar(App.path & "\data\options.ini", "Options", "BuySound")
    Options.SellSound = GetVar(App.path & "\data\options.ini", "Options", "SellSound")
    Options.DeflectAnimation = GetVar(App.path & "\data\options.ini", "Options", "DeflectAnimation")
    Options.CriticalAnimation = GetVar(App.path & "\data\options.ini", "Options", "CriticalAnimation")
    Options.DodgeAnimation = GetVar(App.path & "\data\options.ini", "Options", "DodgeAnimation")
    Options.MaxLevel = GetVar(App.path & "\data\options.ini", "Options", "MaxLevel")
    Options.StatsLevel = GetVar(App.path & "\data\options.ini", "Options", "StatsLevel")
    Options.MaxStat = GetVar(App.path & "\data\options.ini", "Options", "MaxStat")
    Options.LevelUpAnimation = GetVar(App.path & "\data\options.ini", "Options", "LevelUpAnimation")
End Sub

Public Sub LoadDataSizes()
    MAX_MAPS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Maps")
    MAX_ITEMS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Items")
    MAX_SHOPS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Shops")
    MAX_ANIMATIONS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Animations")
    MAX_CLASSES = GetVar(App.path & "\data\options.ini", "Data Sizes", "Classes")
    MAX_EMOTICONS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Emoticons")
    MAX_MORALS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Morals")
    MAX_NPCS = GetVar(App.path & "\data\options.ini", "Data Sizes", "NPCs")
    MAX_QUESTS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Quests")
    MAX_RESOURCES = GetVar(App.path & "\data\options.ini", "Data Sizes", "Resources")
    MAX_SPELLS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Spells")
    MAX_TITLES = GetVar(App.path & "\data\options.ini", "Data Sizes", "Titles")
    MAX_BANS = GetVar(App.path & "\data\options.ini", "Data Sizes", "Bans")
    
    redimData
End Sub

Public Sub redimData()
    ReDim Map(MAX_MAPS)
    ReDim MapBlocks(MAX_MAPS)
    ReDim PlayersOnMap(MAX_MAPS)
    ReDim MapCache(MAX_MAPS)
    ReDim TempEventMap(MAX_MAPS)
    ReDim Item(MAX_ITEMS)
    ReDim MapItem(MAX_MAPS, MAX_MAP_ITEMS)
    ReDim Class(MAX_CLASSES)
    ReDim Animation(MAX_ANIMATIONS)
    ReDim Emoticon(MAX_EMOTICONS)
    ReDim Moral(MAX_MORALS)
    ReDim NPC(MAX_NPCS)
    ReDim Quest(MAX_QUESTS)
    ReDim Resource(MAX_RESOURCES)
    ReDim ResourceCache(MAX_MAPS)
    ReDim Spell(MAX_SPELLS)
    ReDim Title(MAX_TITLES)
    ReDim Ban(MAX_BANS)
    ReDim MapNPC(MAX_MAPS)
    ReDim Shop(MAX_SHOPS)
End Sub

Public Sub SaveDataSizes()
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Maps", Trim$(MAX_MAPS))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Items", Trim$(MAX_ITEMS))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Animations", Trim$(MAX_ANIMATIONS))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Classes", Trim$(MAX_CLASSES))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Emoticons", Trim$(MAX_EMOTICONS))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Morals", Trim$(MAX_MORALS))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "NPCs", Trim$(MAX_NPCS))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Quests", Trim$(MAX_QUESTS))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Resources", Trim$(MAX_RESOURCES))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Spells", Trim$(MAX_SPELLS))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Titles", Trim$(MAX_TITLES))
    Call PutVar(App.path & "\data\options.ini", "Data Sizes", "Bans", Trim$(MAX_BANS))
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As String, ByVal Reason As String)
    Dim IP As String
    Dim i As Integer
    Dim n As Integer

    ' Cut off last portion of IP
    IP = GetPlayerIP(BanPlayerIndex)
    
    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then Exit For
    Next i

    IP = Mid$(IP, 1, i)

    For n = 1 To MAX_BANS
        If Not Len(Trim$(Ban(n).PlayerLogin)) > 0 And Not Len(Trim$(Ban(n).playerName)) > 0 Then
            With Ban(n)
                .Date = Date
                
                If BannedByIndex <> "server" Then
                    .By = GetPlayerName(BannedByIndex)
                Else
                    .By = "server"
                End If
                
                .Time = Time
                .HDSerial = GetPlayerHDSerial(BanPlayerIndex)
                .IP = IP
                .PlayerLogin = GetPlayerLogin(BanPlayerIndex)
                .playerName = GetPlayerName(BanPlayerIndex)
                .Reason = Reason
            End With
            Call SaveBan(n)
            Exit For
        End If
    Next n

    If Not BannedByIndex = "server" Then
        If Len(Reason) Then
            AdminMsg GetPlayerName(BanPlayerIndex) & " has been banned by " & GetPlayerName(BannedByIndex) & " for " & Reason & "!", BrightBlue
            AddLog GetPlayerName(BannedByIndex) & "/" & GetPlayerIP(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & "/" & GetPlayerIP(BanPlayerIndex) & " for " & Reason & ".", "Bans"
            AlertMsg BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & " for " & Reason & "!"
        Else
            AdminMsg GetPlayerName(BanPlayerIndex) & " has been banned by " & GetPlayerName(BannedByIndex) & "!", BrightBlue
            AddLog GetPlayerName(BannedByIndex) & "/" & GetPlayerIP(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & "/" & GetPlayerIP(BanPlayerIndex) & ".", "Admin"
            AlertMsg BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!"
        End If
    Else
        AdminMsg GetPlayerName(BanPlayerIndex) & " has been banned by the server!", BrightBlue
        AddLog GetPlayerName(BanPlayerIndex) & "/" & GetPlayerIP(BanPlayerIndex) & " was banned by the server!", "Admin"
        AlertMsg BanPlayerIndex, "You have been banned by the server!"
    End If
    Call LeftGame(BanPlayerIndex)
End Sub

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    
    Call ChkDir(App.path & "\data\accounts\", Trim$(Name))
    filename = "\data\accounts\" & Trim$(Name) & "\data.bin"

    If FileExist(filename) Then
        AccountExist = True
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim filename As String
    Dim RightPassword As String * NAME_LENGTH
    Dim nFileNum As Long

    PasswordOK = False

    If AccountExist(Name) Then
        filename = App.path & "\data\accounts\" & Trim$(Name) & "\data.bin"
        nFileNum = FreeFile
        Open filename For Binary As #nFileNum
        Get #nFileNum, NAME_LENGTH, RightPassword
        Close #nFileNum
       
        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If
End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearAccount index
    
    Account(index).Login = Name
    Account(index).Password = Password
    
    Call SaveAccount(index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Dim charLogin() As String
    
    Call FileCopy(App.path & "\data\accounts\charlist.txt", App.path & "\data\accounts\chartemp.txt")
    
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.path & "\data\accounts\chartemp.txt" For Input As #f1
    
    f2 = FreeFile
    Open App.path & "\data\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s
        charLogin = Split(s, ":") ' Character Editor
        If Trim$(LCase$(charLogin(0))) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If
    Loop

    Close #f1
    Close #f2
    Call Kill(App.path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long) As Boolean
    If Len(Trim$(Account(index).Chars(GetPlayerChar(index)).Name)) > 0 And Len(Trim$(Account(index).Chars(GetPlayerChar(index)).Name)) <= NAME_LENGTH Then
        CharExist = True
    End If
End Function

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Gender As Byte, ByVal ClassNum As Byte)
    Dim i As Long, F As Long

    With Account(index).Chars(GetPlayerChar(index))
        ' Basic things
        .Name = Name
        .Gender = Gender
        .Class = ClassNum
        
        ' Sprite and face
        If .Gender = GENDER_MALE Then
            .Sprite = Class(ClassNum).MaleSprite
            .Face = Class(ClassNum).MaleFace
        Else
            .Sprite = Class(ClassNum).FemaleSprite
            .Face = Class(ClassNum).FemaleFace
        End If
    
        ' Level
        .Level = 1
    
        ' Stats
        For i = 1 To Stats.Stat_count - 1
            .Stat(i) = Class(ClassNum).Stat(i)
        Next
        
        ' Skills
        For i = 1 To Skills.Skill_Count - 1
            Call SetPlayerSkill(index, 1, i)
        Next
        
        ' Set the player's start values
        .Dir = Class(GetPlayerClass(index)).Dir
        .Map = Class(GetPlayerClass(index)).Map
        .x = Class(GetPlayerClass(index)).x
        .Y = Class(GetPlayerClass(index)).Y
        
        ' Vitals
        .Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        .Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)
        
        ' Restore vitals
        Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
        Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))
        
        ' Set the checkpoint values
        .CheckPointMap = .Map
        .CheckPointX = .x
        .CheckPointY = .Y
        
        ' Set the trade status value
        .CanTrade = True
    
        ' Set the status to nothing
        .Status = vbNullString
        
        ' Check for new title
        Call CheckPlayerNewTitle(index, False)
        
        ' Set starter equipment
        For i = 1 To MAX_INV
            If Class(ClassNum).StartItem(i) > 0 Then
                ' Item exist?
                If Len(Trim$(Item(Class(ClassNum).StartItem(i)).Name)) > 0 Then
                    .Inv(i).Num = Class(ClassNum).StartItem(i)
                    .Inv(i).Value = Class(ClassNum).StartItemValue(i)
                End If
            End If
        Next
        
        ' Set start spells
        For i = 1 To MAX_PLAYER_SPELLS
            If Class(ClassNum).StartSpell(i) > 0 Then
                ' Spell exist?
                If Len(Trim$(Spell(Class(ClassNum).StartSpell(i)).Name)) > 0 Then
                    .Spell(i) = Class(ClassNum).StartSpell(i)
                End If
            End If
        Next
        
        ReDim Preserve .QuestCompleted(MAX_QUESTS)
        ReDim Preserve .QuestCLI(MAX_QUESTS)
        ReDim Preserve .QuestAmount(MAX_QUESTS)
        ReDim Preserve .QuestTask(MAX_QUESTS)
        
        For i = 1 To MAX_QUESTS
            ReDim Preserve .QuestAmount(i).ID(MAX_NPCS)
        Next
        
    End With
    
    ' Append name to file
    F = FreeFile
    
    Open App.path & "\data\accounts\charlist.txt" For Append As #F
        Print #F, Name & ":" & Account(index).Login ' Character Editor
    Close #F
    
    Call SaveAccount(index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    Dim charLogin() As String
    
    ' Check if the master charlist file exists for checking duplicate names, and if it doesnt make it
    If Not FileExist("\data\accounts\charlist.txt") Then
        F = FreeFile
        Open App.path & "\data\accounts\charlist.txt" For Output As #F
        Close #F
    End If

    F = FreeFile
    
    Open App.path & "\data\accounts\charlist.txt" For Input As #F
        Do While Not EOF(F)
            Input #F, s
            charLogin = Split(s, ":") ' Character Editor
            If Trim$(LCase$(charLogin(0))) = Trim$(LCase$(Name)) Then
                FindChar = True
                Close #F
                Exit Function
            End If
        Loop
    Close #F
End Function

' Character Editor
Function GetCharList() As String
    Dim F As Long, counter As Long
    Dim s As String, total As String
    Dim charLogin() As String
    
    F = FreeFile
    
    Open App.path & "\data\accounts\charlist.txt" For Input As #F
        Do While Not EOF(F)
            Input #F, s
            charLogin = Split(s, ":")
            counter = counter + 1
            total = total & charLogin(0) & ","
        Loop
    Close #F
    
    If counter > 0 Then
        total = Left$(total, Len(total) - 1)
    End If
    GetCharList = total
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SaveAccount(i)
        End If
    Next
End Sub

Sub SaveAccount(ByVal index As Long)
    Dim filename As String
    Dim F As Long

    Call ChkDir(App.path & "\data\accounts\", GetPlayerLogin(index))
    filename = App.path & "\data\accounts\" & GetPlayerLogin(index) & "\data.bin"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Account(index)
    Close #F
End Sub

Sub loadAccount(ByVal index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long, i As Long
    Dim Length As Long

    ClearAccount index
    
    filename = App.path & "\data\accounts\" & Trim$(Name) & "\data.bin"
    F = FreeFile
    
    Open filename For Binary As #F
        Get #F, , Account(index)
    Close #F
End Sub

Sub ClearAccount(ByVal index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(tempplayer(index)), LenB(tempplayer(index)))
    tempplayer(index).HDSerial = vbNullString
    Set tempplayer(index).buffer = New clsBuffer
    
    ZeroMemory ByVal VarPtr(Account(index)), LenB(Account(index))
    Account(index).Login = vbNullString
    Account(index).Password = vbNullString
    Account(index).CurrentChar = 1
    Account(index).Chars(GetPlayerChar(index)).Name = vbNullString
    Account(index).Chars(GetPlayerChar(index)).Status = vbNullString
    Account(index).Chars(GetPlayerChar(index)).Class = 1
    
    For i = 1 To Stats.Stat_count - 1
        Call SetPlayerStat(index, 1, i)
    Next
    
    For i = 1 To Skills.Skill_Count - 1
        Call SetPlayerSkill(index, 1, i)
    Next
    
    frmServer.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(3) = vbNullString
End Sub

' ***********
' ** Classes **
' ***********
Sub SaveClasses()
    Dim i As Long

    For i = 1 To MAX_CLASSES
        Call SaveClass(i)
    Next
End Sub

Sub SaveClass(ByVal ClassNum As Long)
    Dim filename As String
    Dim F  As Long
    
    filename = App.path & "\data\classes\" & ClassNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Class(ClassNum)
    Close #F
End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    Call CheckClasses

    For i = 1 To MAX_CLASSES
        filename = App.path & "\data\classes\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Class(i)
        Close #F
    Next
End Sub

Sub CheckClasses()
    Dim i As Long

    For i = 1 To MAX_CLASSES
        If Not FileExist("\data\classes\" & i & ".dat") Then
            Call ClearClass(i)
            Call SaveClass(i)
        End If
    Next
End Sub

Sub ClearClass(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Class(index)), LenB(Class(index)))
    Class(index).Name = vbNullString
    Class(index).CombatTree = 1
    Class(index).Map = 1
    Class(index).Color = 15
End Sub

Sub ClearClasses()
    Dim i As Long

    For i = 1 To MAX_CLASSES
        Call ClearClass(i)
    Next
End Sub

' ***********
' ** Items **
' ***********
Sub SaveItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next
End Sub

Sub SaveItem(ByVal ItemNum As Integer)
    Dim filename As String
    Dim F  As Long
    
    filename = App.path & "\data\items\" & ItemNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Item(ItemNum)
    Close #F
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.path & "\data\items\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Item(i)
        Close #F
    Next
End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        If Not FileExist("\data\items\" & i & ".dat") Then
            Call ClearItem(i)
            Call SaveItem(i)
        End If
    Next
End Sub

Sub ClearItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = vbNullString
    Item(index).Rarity = 1
End Sub

Sub ClearItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS
        Call ClearItem(i)
    Next
End Sub

' ***********
' ** Shops **
' ***********
Sub SaveShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next
End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.path & "\data\shops\" & ShopNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Shop(ShopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.path & "\data\shops\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Shop(i)
        Close #F
    Next
End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        If Not FileExist("\data\shops\" & i & ".dat") Then
            Call ClearShop(i)
            Call SaveShop(i)
        End If
    Next
End Sub

Sub ClearShop(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Shop(index)), LenB(Shop(index)))
    Shop(index).Name = vbNullString
End Sub

Sub ClearShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS
        Call ClearShop(i)
    Next
End Sub

' ************
' ** Spells **
' ************
Sub SaveSpell(ByVal SpellNum As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.path & "\data\spells\" & SpellNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Spell(SpellNum)
    Close #F
End Sub

Sub SaveSpells()
    Dim i As Long
    
    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next
End Sub

Sub LoadSpells()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    Call CheckSpells

    For i = 1 To MAX_SPELLS
        filename = App.path & "\data\spells\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Spell(i)
        Close #F
    Next
End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        If Not FileExist("\data\spells\" & i & ".dat") Then
            Call ClearSpell(i)
            Call SaveSpell(i)
        End If
    Next
End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).Desc = vbNullString
    Spell(index).LevelReq = 1 ' Needs to be 1 for the spell editor
    Spell(index).Sound = vbNullString
End Sub

Sub ClearSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS
        Call ClearSpell(i)
    Next
End Sub

' **********
' ** NPCs **
' **********
Sub SaveNPCs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNPC(i)
    Next
End Sub

Sub SaveNPC(ByVal NPCNum As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.path & "\data\npcs\" & NPCNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , NPC(NPCNum)
    Close #F
End Sub

Sub LoadNPCs()
    Dim i As Long

    Call CheckNPCs

    For i = 1 To MAX_NPCS
        Call LoadNPC(i)
    Next
End Sub

Sub LoadNPC(NPCNum As Long)
    Dim F As Long
    Dim filename As String
    
    filename = App.path & "\data\npcs\" & NPCNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Get #F, , NPC(NPCNum)
    Close #F
End Sub

Sub CheckNPCs()
    Dim i As Integer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    
    For i = 1 To MAX_NPCS
        If Not FileExist("\data\npcs\" & i & ".dat") Then
            Call ClearNPC(i)
            Call SaveNPC(i)
        End If
    Next
End Sub

Sub ClearNPC(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(NPC(index)), LenB(NPC(index)))
    NPC(index).Name = vbNullString
    NPC(index).Title = vbNullString
    NPC(index).AttackSay = vbNullString
    NPC(index).Music = vbNullString
    NPC(index).Sound = vbNullString
End Sub

Sub ClearNPCs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNPC(i)
    Next
End Sub

' ***************
' ** Resources **
' ***************
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next
End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.path & "\data\resources\" & ResourceNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Integer
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.path & "\data\resources\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next
End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\data\resources\" & i & ".dat") Then
            Call ClearResource(i)
            Call SaveResource(i)
        End If
    Next
End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).FailMessage = vbNullString
    Resource(index).Sound = vbNullString
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' ****************
' ** Animations **
' ****************
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next
End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.path & "\data\animations\" & AnimationNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Integer
    Dim F As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.path & "\data\animations\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next
End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If Not FileExist("\data\animations\" & i & ".dat") Then
            Call ClearAnimation(i)
            Call SaveAnimation(i)
        End If
    Next
End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).Name = vbNullString
    Animation(index).Sound = vbNullString
End Sub

Sub ClearAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call ClearAnimation(i)
    Next
End Sub

' **********
' ** Maps **
' **********
Sub SaveMap(ByVal MapNum As Long)
    Dim filename As String
    Dim F As Long
    Dim x As Long
    Dim Y As Long, i As Long, z As Long, w As Long
    
    filename = App.path & "\data\maps\" & MapNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Map(MapNum).Name
        Put #F, , Map(MapNum).Music
        Put #F, , Map(MapNum).BGS
        Put #F, , Map(MapNum).Revision
        Put #F, , Map(MapNum).Moral
        Put #F, , Map(MapNum).Up
        Put #F, , Map(MapNum).Down
        Put #F, , Map(MapNum).Left
        Put #F, , Map(MapNum).Right
        Put #F, , Map(MapNum).BootMap
        Put #F, , Map(MapNum).BootX
        Put #F, , Map(MapNum).BootY
        
        Put #F, , Map(MapNum).Weather
        Put #F, , Map(MapNum).WeatherIntensity
        
        Put #F, , Map(MapNum).Fog
        Put #F, , Map(MapNum).FogSpeed
        Put #F, , Map(MapNum).FogOpacity
        
        Put #F, , Map(MapNum).Panorama
        
        Put #F, , Map(MapNum).Red
        Put #F, , Map(MapNum).Green
        Put #F, , Map(MapNum).Blue
        Put #F, , Map(MapNum).Alpha
        
        Put #F, , Map(MapNum).MaxX
        Put #F, , Map(MapNum).MaxY
        
        Put #F, , Map(MapNum).NPC_HighIndex
    
        For x = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                Put #F, , Map(MapNum).Tile(x, Y)
            Next
        Next
    
        For x = 1 To MAX_MAP_NPCS
            Put #F, , Map(MapNum).NPC(x)
            Put #F, , Map(MapNum).NPCSpawnType(x)
        Next
    Close #F
    
    ' This is for event saving, it is in .ini files becuase there are non-limited values (strings) that Can't easily be loaded/saved in the normal manner.
    filename = App.path & "\data\maps\" & MapNum & "_eventdata.dat"
    PutVar filename, "Events", "EventCount", Val(Map(MapNum).EventCount)
    
    If Map(MapNum).EventCount > 0 Then
        For i = 1 To Map(MapNum).EventCount
            With Map(MapNum).Events(i)
                PutVar filename, "Event" & i, "Name", .Name
                PutVar filename, "Event" & i, "Global", Val(.Global)
                PutVar filename, "Event" & i, "x", Val(.x)
                PutVar filename, "Event" & i, "y", Val(.Y)
                PutVar filename, "Event" & i, "PageCount", Val(.PageCount)
            End With
            
            If Map(MapNum).Events(i).PageCount > 0 Then
                For x = 1 To Map(MapNum).Events(i).PageCount
                    With Map(MapNum).Events(i).Pages(x)
                        PutVar filename, "Event" & i & "Page" & x, "chkVariable", Val(.chkVariable)
                        PutVar filename, "Event" & i & "Page" & x, "VariableIndex", Val(.VariableIndex)
                        PutVar filename, "Event" & i & "Page" & x, "VariableCondition", Val(.VariableCondition)
                        PutVar filename, "Event" & i & "Page" & x, "VariableCompare", Val(.VariableCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkSwitch", Val(.chkSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "SwitchIndex", Val(.SwitchIndex)
                        PutVar filename, "Event" & i & "Page" & x, "SwitchCompare", Val(.SwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkHasItem", Val(.chkHasItem)
                        PutVar filename, "Event" & i & "Page" & x, "HasItemIndex", Val(.HasItemIndex)
                        
                        PutVar filename, "Event" & i & "Page" & x, "chkSelfSwitch", Val(.chkSelfSwitch)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchIndex", Val(.SelfSwitchIndex)
                        PutVar filename, "Event" & i & "Page" & x, "SelfSwitchCompare", Val(.SelfSwitchCompare)
                        
                        PutVar filename, "Event" & i & "Page" & x, "GraphicType", Val(.GraphicType)
                        PutVar filename, "Event" & i & "Page" & x, "Graphic", Val(.Graphic)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX", Val(.GraphicX)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY", Val(.GraphicY)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicX2", Val(.GraphicX2)
                        PutVar filename, "Event" & i & "Page" & x, "GraphicY2", Val(.GraphicY2)
                        
                        PutVar filename, "Event" & i & "Page" & x, "MoveType", Val(.MoveType)
                        PutVar filename, "Event" & i & "Page" & x, "MoveSpeed", Val(.MoveSpeed)
                        PutVar filename, "Event" & i & "Page" & x, "MoveFreq", Val(.MoveFreq)
                        
                        PutVar filename, "Event" & i & "Page" & x, "IgnoreMoveRoute", Val(.IgnoreMoveRoute)
                        PutVar filename, "Event" & i & "Page" & x, "RepeatMoveRoute", Val(.RepeatMoveRoute)
                        
                        PutVar filename, "Event" & i & "Page" & x, "MoveRouteCount", Val(.MoveRouteCount)
                        
                        If .MoveRouteCount > 0 Then
                            For Y = 1 To .MoveRouteCount
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Index", Val(.MoveRoute(Y).index)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data1", Val(.MoveRoute(Y).Data1)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data2", Val(.MoveRoute(Y).Data2)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data3", Val(.MoveRoute(Y).Data3)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data4", Val(.MoveRoute(Y).Data4)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data5", Val(.MoveRoute(Y).Data5)
                                PutVar filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data6", Val(.MoveRoute(Y).Data6)
                            Next
                        End If
                        
                        PutVar filename, "Event" & i & "Page" & x, "WalkAnim", Val(.WalkAnim)
                        PutVar filename, "Event" & i & "Page" & x, "DirFix", Val(.DirFix)
                        PutVar filename, "Event" & i & "Page" & x, "WalkThrough", Val(.WalkThrough)
                        PutVar filename, "Event" & i & "Page" & x, "ShowName", Val(.ShowName)
                        PutVar filename, "Event" & i & "Page" & x, "Trigger", Val(.Trigger)
                        PutVar filename, "Event" & i & "Page" & x, "CommandListCount", Val(.CommandListCount)
                        
                        PutVar filename, "Event" & i & "Page" & x, "Position", Val(.Position)
                    End With
                    
                    If Map(MapNum).Events(i).Pages(x).CommandListCount > 0 Then
                        For Y = 1 To Map(MapNum).Events(i).Pages(x).CommandListCount
                            PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "CommandCount", Val(Map(MapNum).Events(i).Pages(x).CommandList(Y).CommandCount)
                            PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "ParentList", Val(Map(MapNum).Events(i).Pages(x).CommandList(Y).ParentList)
                            If Map(MapNum).Events(i).Pages(x).CommandList(Y).CommandCount > 0 Then
                                For z = 1 To Map(MapNum).Events(i).Pages(x).CommandList(Y).CommandCount
                                    With Map(MapNum).Events(i).Pages(x).CommandList(Y).Commands(z)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Index", Val(.index)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Text1", .Text1
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Text2", .Text2
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Text3", .Text3
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Text4", .Text4
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Text5", .Text5
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Data1", Val(.Data1)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Data2", Val(.Data2)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Data3", Val(.Data3)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Data4", Val(.Data4)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Data5", Val(.Data5)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "Data6", Val(.Data6)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "ConditionalBranchCommandList", Val(.ConditionalBranch.CommandList)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "ConditionalBranchCondition", Val(.ConditionalBranch.Condition)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "ConditionalBranchData1", Val(.ConditionalBranch.Data1)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "ConditionalBranchData2", Val(.ConditionalBranch.Data2)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "ConditionalBranchData3", Val(.ConditionalBranch.Data3)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "ConditionalBranchElseCommandList", Val(.ConditionalBranch.ElseCommandList)
                                        PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "MoveRouteCount", Val(.MoveRouteCount)
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Index", Val(.MoveRoute(w).index)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data1", Val(.MoveRoute(w).Data1)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data2", Val(.MoveRoute(w).Data2)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data3", Val(.MoveRoute(w).Data3)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data4", Val(.MoveRoute(w).Data4)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data5", Val(.MoveRoute(w).Data5)
                                                PutVar filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & z & "MoveRoute" & w & "Data6", Val(.MoveRoute(w).Data6)
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
        
    DoEvents
End Sub

Sub SaveMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next
End Sub

Sub LoadMaps()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim x As Long
    Dim Y As Long, z As Long, p As Long, w As Long
    Dim newtileset As Long, newtiley As Long
  
    Call CheckMaps
            
    For i = 1 To MAX_MAPS
        On Error Resume Next
        filename = App.path & "\data\maps\" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).BGS
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        
        Get #F, , Map(i).Weather
        Get #F, , Map(i).WeatherIntensity
        
        Get #F, , Map(i).Fog
        Get #F, , Map(i).FogSpeed
        Get #F, , Map(i).FogOpacity
        
        Get #F, , Map(i).Panorama
        
        Get #F, , Map(i).Red
        Get #F, , Map(i).Green
        Get #F, , Map(i).Blue
        Get #F, , Map(i).Alpha
        
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        Get #F, , Map(i).NPC_HighIndex
        
        For x = 0 To Map(i).MaxX
            For Y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(x, Y)
            Next
        Next
        
        For x = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).NPC(x)
            Get #F, , Map(i).NPCSpawnType(x)
            MapNPC(i).NPC(x).Num = Map(i).NPC(x)
        Next

        Close #F
        
        CacheResources i
        DoEvents
        CacheMapBlocks i
    Next
    
    For z = 1 To MAX_MAPS
        filename = App.path & "\data\maps\" & z & "_eventdata.dat"
        Map(z).EventCount = Val(GetVar(filename, "Events", "EventCount"))
        
        If Map(z).EventCount > 0 Then
            ReDim Map(z).Events(0 To Map(z).EventCount)
            For i = 1 To Map(z).EventCount
                With Map(z).Events(i)
                    .Name = GetVar(filename, "Event" & i, "Name")
                    .Global = Val(GetVar(filename, "Event" & i, "Global"))
                    .x = Val(GetVar(filename, "Event" & i, "x"))
                    .Y = Val(GetVar(filename, "Event" & i, "y"))
                    .PageCount = Val(GetVar(filename, "Event" & i, "PageCount"))
                End With
                If Map(z).Events(i).PageCount > 0 Then
                    ReDim Map(z).Events(i).Pages(0 To Map(z).Events(i).PageCount)
                    For x = 1 To Map(z).Events(i).PageCount
                        With Map(z).Events(i).Pages(x)
                            .chkVariable = Val(GetVar(filename, "Event" & i & "Page" & x, "chkVariable"))
                            .VariableIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableIndex"))
                            .VariableCondition = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableCondition"))
                            .VariableCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "VariableCompare"))
                            
                            .chkSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSwitch"))
                            .SwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "SwitchIndex"))
                            .SwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "SwitchCompare"))
                            
                            .chkHasItem = Val(GetVar(filename, "Event" & i & "Page" & x, "chkHasItem"))
                            .HasItemIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "HasItemIndex"))
                            
                            .chkSelfSwitch = Val(GetVar(filename, "Event" & i & "Page" & x, "chkSelfSwitch"))
                            .SelfSwitchIndex = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchIndex"))
                            .SelfSwitchCompare = Val(GetVar(filename, "Event" & i & "Page" & x, "SelfSwitchCompare"))
                            
                            .GraphicType = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicType"))
                            .Graphic = Val(GetVar(filename, "Event" & i & "Page" & x, "Graphic"))
                            .GraphicX = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX"))
                            .GraphicY = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY"))
                            .GraphicX2 = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicX2"))
                            .GraphicY2 = Val(GetVar(filename, "Event" & i & "Page" & x, "GraphicY2"))
                            
                            .MoveType = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveType"))
                            .MoveSpeed = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveSpeed"))
                            .MoveFreq = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveFreq"))
                            
                            .IgnoreMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & x, "IgnoreMoveRoute"))
                            .RepeatMoveRoute = Val(GetVar(filename, "Event" & i & "Page" & x, "RepeatMoveRoute"))
                            
                            .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRouteCount"))
                            
                            If .MoveRouteCount > 0 Then
                                ReDim Map(z).Events(i).Pages(x).MoveRoute(0 To .MoveRouteCount)
                                For Y = 1 To .MoveRouteCount
                                    .MoveRoute(Y).index = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Index"))
                                    .MoveRoute(Y).Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data1"))
                                    .MoveRoute(Y).Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data2"))
                                    .MoveRoute(Y).Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data3"))
                                    .MoveRoute(Y).Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data4"))
                                    .MoveRoute(Y).Data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data5"))
                                    .MoveRoute(Y).Data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "MoveRoute" & Y & "Data6"))
                                Next
                            End If
                            
                            .WalkAnim = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkAnim"))
                            .DirFix = Val(GetVar(filename, "Event" & i & "Page" & x, "DirFix"))
                            .WalkThrough = Val(GetVar(filename, "Event" & i & "Page" & x, "WalkThrough"))
                            .ShowName = Val(GetVar(filename, "Event" & i & "Page" & x, "ShowName"))
                            .Trigger = Val(GetVar(filename, "Event" & i & "Page" & x, "Trigger"))
                            .CommandListCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandListCount"))
                         
                            .Position = Val(GetVar(filename, "Event" & i & "Page" & x, "Position"))
                        End With
                            
                        If Map(z).Events(i).Pages(x).CommandListCount > 0 Then
                            ReDim Map(z).Events(i).Pages(x).CommandList(0 To Map(z).Events(i).Pages(x).CommandListCount)
                            For Y = 1 To Map(z).Events(i).Pages(x).CommandListCount
                                Map(z).Events(i).Pages(x).CommandList(Y).CommandCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "CommandCount"))
                                Map(z).Events(i).Pages(x).CommandList(Y).ParentList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "ParentList"))
                                If Map(z).Events(i).Pages(x).CommandList(Y).CommandCount > 0 Then
                                    ReDim Map(z).Events(i).Pages(x).CommandList(Y).Commands(Map(z).Events(i).Pages(x).CommandList(Y).CommandCount)
                                    For p = 1 To Map(z).Events(i).Pages(x).CommandList(Y).CommandCount
                                        With Map(z).Events(i).Pages(x).CommandList(Y).Commands(p)
                                            .index = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Index"))
                                            .Text1 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Text1")
                                            .Text2 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Text2")
                                            .Text3 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Text3")
                                            .Text4 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Text4")
                                            .Text5 = GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Text5")
                                            .Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Data1"))
                                            .Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Data2"))
                                            .Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Data3"))
                                            .Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Data4"))
                                            .Data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Data5"))
                                            .Data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "Data6"))
                                            .ConditionalBranch.CommandList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "ConditionalBranchCommandList"))
                                            .ConditionalBranch.Condition = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "ConditionalBranchCondition"))
                                            .ConditionalBranch.Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "ConditionalBranchData1"))
                                            .ConditionalBranch.Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "ConditionalBranchData2"))
                                            .ConditionalBranch.Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "ConditionalBranchData3"))
                                            .ConditionalBranch.ElseCommandList = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "ConditionalBranchElseCommandList"))
                                            .MoveRouteCount = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "MoveRouteCount"))
                                            If .MoveRouteCount > 0 Then
                                                ReDim .MoveRoute(1 To .MoveRouteCount)
                                                For w = 1 To .MoveRouteCount
                                                    .MoveRoute(w).index = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Index"))
                                                    .MoveRoute(w).Data1 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data1"))
                                                    .MoveRoute(w).Data2 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data2"))
                                                    .MoveRoute(w).Data3 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data3"))
                                                    .MoveRoute(w).Data4 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data4"))
                                                    .MoveRoute(w).Data5 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data5"))
                                                    .MoveRoute(w).Data6 = Val(GetVar(filename, "Event" & i & "Page" & x, "CommandList" & Y & "Command" & p & "MoveRoute" & w & "Data6"))
                                                Next
                                            End If
                                        End With
                                    Next
                                End If
                            Next
                        End If
                    Next
                End If
            Next
        End If
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        If Not FileExist("\data\maps\" & i & ".dat") Then
            Call ClearMap(i)
            Call SaveMap(i)
        End If
    Next
End Sub

Sub ClearMapItem(ByVal index As Long, ByVal MapNum As Integer)
    Call ZeroMemory(ByVal VarPtr(MapItem(MapNum, index)), LenB(MapItem(MapNum, index)))
    MapItem(MapNum, index).playerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, Y)
        Next
    Next
End Sub

Sub ClearMapNPC(ByVal index As Long, ByVal MapNum As Integer)
    Call ZeroMemory(ByVal VarPtr(MapNPC(MapNum).NPC(index)), LenB(MapNPC(MapNum).NPC(index)))
End Sub

Sub ClearMapNPCs()
    Dim x As Long
    Dim Y As Long
    
    For Y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNPC(x, Y)
        Next
    Next
End Sub

Sub ClearMap(ByVal MapNum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(MapNum)), LenB(Map(MapNum)))
    
    Map(MapNum).Name = vbNullString
    Map(MapNum).Music = vbNullString
    Map(MapNum).BGS = vbNullString
    Map(MapNum).Moral = 1
    Map(MapNum).MaxX = MIN_MAPX
    Map(MapNum).MaxY = MIN_MAPY
    
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)
    
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO
    
    ' Reset the map cache array for this map.
    MapCache(MapNum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next
End Sub

' ************
' ** Guilds **
' ************
Sub SaveGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        Call SaveGuild(i)
    Next
End Sub

Sub SaveGuild(ByVal GuildNum As Long)
    Dim filename As String
    Dim F  As Long
    
    filename = App.path & "\data\guilds\" & GuildNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Guild(GuildNum)
    Close #F
End Sub

Sub LoadGuilds()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    Call CheckGuild

    For i = 1 To MAX_GUILDS
        filename = App.path & "\data\guilds\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Guild(i)
        Close #F
    Next
End Sub

Sub CheckGuild()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        If Not FileExist("\data\guilds\" & i & ".dat") Then
            Call ClearGuild(i)
            Call SaveGuild(i)
        End If
    Next
End Sub

Sub ClearGuilds()
    Dim i As Long

    For i = 1 To MAX_GUILDS
        Call ClearGuild(i)
    Next
End Sub

Sub ClearGuild(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Guild(index)), LenB(Guild(index)))
    Guild(index).Name = vbNullString
    Guild(index).MOTD = vbNullString
End Sub

' ************
' ** Bans **
' ************
Sub SaveBan(ByVal BanNum As Long)
    Dim F As Long
    Dim filename As String
    
    F = FreeFile
    filename = App.path & "\data\bans\" & BanNum & ".dat"
    
    Open filename For Binary As #F
        Put #F, , Ban(BanNum)
    Close #F
End Sub

Sub CheckBans()
    Dim i As Long

    For i = 1 To MAX_BANS
        If Not FileExist("\data\bans\" & i & ".dat") Then
            Call ClearBan(i)
            Call SaveBan(i)
        End If
    Next
End Sub

Sub ClearBans()
    Dim i As Long
    
    For i = 1 To MAX_BANS
        Call ClearBan(i)
    Next
End Sub

Sub ClearBan(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Ban(index)), LenB(Ban(index)))
    Ban(index).PlayerLogin = vbNullString
    Ban(index).playerName = vbNullString
    Ban(index).Reason = vbNullString
    Ban(index).IP = vbNullString
    Ban(index).HDSerial = vbNullString
    Ban(index).Time = vbNullString
    Ban(index).By = vbNullString
    Ban(index).Date = vbNullString
End Sub

' ************
' ** Titles **
' ************
Sub SaveTitle(ByVal TitleNum As Long)
    Dim F As Long
    Dim filename As String

    F = FreeFile
    filename = App.path & "\data\titles\" & TitleNum & ".dat"
    
    Open filename For Binary As #F
        Put #F, , Title(TitleNum)
    Close #F
End Sub

Sub LoadTitles()
    Dim i As Long

    CheckTitles
    
    For i = 1 To MAX_TITLES
        Call LoadTitle(i)
    Next
End Sub

Sub LoadTitle(index As Long)
    Dim F As Long
    Dim filename  As String

    F = FreeFile
    filename = App.path & "\data\titles\" & index & ".dat"
    
    Open filename For Binary As #F
        Get #F, , Title(index)
    Close #F
End Sub

Sub CheckTitles()
    Dim i As Long

    For i = 1 To MAX_TITLES
        If Not FileExist("\data\titles\" & i & ".dat") Then
            Call ClearTitle(i)
            Call SaveTitle(i)
        End If
    Next
End Sub

Sub ClearTitles()
    Dim i As Long
    
    For i = 1 To MAX_TITLES
        Call ClearTitle(i)
    Next
End Sub

Sub ClearTitle(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Title(index)), LenB(Title(index)))
    Title(index).Name = vbNullString
End Sub

' ************
' ** Morals **
' ************
Sub SaveMorals()
    Dim i As Long

    For i = 1 To MAX_MORALS
        Call SaveMoral(i)
    Next
End Sub

Sub SaveMoral(ByVal MoralNum As Long)
    Dim filename As String
    Dim F  As Long
    
    filename = App.path & "\data\morals\" & MoralNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Moral(MoralNum)
    Close #F
End Sub

Sub LoadMorals()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    Call CheckMorals

    For i = 1 To MAX_MORALS
        filename = App.path & "\data\morals\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Moral(i)
        Close #F
    Next
End Sub

Sub CheckMorals()
    Dim i As Long

    For i = 1 To MAX_MORALS
        If Not FileExist("\data\morals\" & i & ".dat") Then
            Call ClearMoral(i)
            Call SaveMoral(i)
        End If
    Next
End Sub

Sub ClearMoral(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Moral(index)), LenB(Moral(index)))
    Moral(index).Name = vbNullString
End Sub

Sub ClearMorals()
    Dim i As Long

    For i = 1 To MAX_MORALS
        Call ClearMoral(i)
    Next
End Sub

' **************
' ** Emoticons **
' **************
Sub SaveEmoticons()
    Dim i As Long

    For i = 1 To MAX_EMOTICONS
        Call SaveEmoticon(i)
    Next
End Sub

Sub SaveEmoticon(ByVal EmoticonNum As Long)
    Dim filename As String
    Dim F  As Long
    
    filename = App.path & "\data\emoticons\" & EmoticonNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , Emoticon(EmoticonNum)
    Close #F
End Sub

Sub LoadEmoticons()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    
    Call CheckEmoticons

    For i = 1 To MAX_EMOTICONS
        filename = App.path & "\data\emoticons\" & i & ".dat"
        F = FreeFile
        
        Open filename For Binary As #F
            Get #F, , Emoticon(i)
        Close #F
    Next
End Sub

Sub CheckEmoticons()
    Dim i As Long

    For i = 1 To MAX_EMOTICONS
        If Not FileExist("\data\emoticons\" & i & ".dat") Then
            Call ClearEmoticon(i)
            Call SaveEmoticon(i)
        End If
    Next
End Sub

Sub ClearEmoticons()
    Dim i As Long

    For i = 1 To MAX_EMOTICONS
        Call ClearEmoticon(i)
    Next
End Sub

Sub ClearEmoticon(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Emoticon(index)), LenB(Emoticon(index)))
    Emoticon(index).Command = "/"
End Sub

' ***********
' ** Party **
' ***********
Sub ClearParty(ByVal PartyNum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(PartyNum)), LenB(Party(PartyNum)))
End Sub

Sub SaveTempGuildMember(ByVal index As Long, ByVal Login As String)
    Dim filename As String
    Dim F As Long

    filename = App.path & "\data\accounts\" & Trim$(Login) & "\data.bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , TempGuildMember(index)
    Close #F
End Sub

Sub LoadTempGuildMember(ByVal index As Long, ByVal Login As String)
    Dim filename As String
    Dim F As Long
    
    Call ClearTempGuildMember(index)
    filename = App.path & "\data\Accounts\" & Trim$(Login) & "\data.bin"
    F = FreeFile
    Open filename For Binary As #F
    Get #F, , TempGuildMember(index)
    Close #F
End Sub

Sub ClearTempGuildMember(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(TempGuildMember(index)), LenB(TempGuildMember(index)))
End Sub

Sub SaveSwitches()
    Dim i As Long, filename As String
    
    filename = App.path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        Call PutVar(filename, "Switches", "Switch" & CStr(i) & "Name", Switches(i))
    Next
End Sub

Sub SaveVariables()
    Dim i As Long, filename As String
    
    filename = App.path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        Call PutVar(filename, "Variables", "Variable" & CStr(i) & "Name", Variables(i))
    Next
End Sub

Sub LoadSwitches()
    Dim i As Long, filename As String
    
    filename = App.path & "\data\switches.ini"
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = GetVar(filename, "Switches", "Switch" & CStr(i) & "Name")
    Next
End Sub

Sub LoadVariables()
    Dim i As Long, filename As String
    
    filename = App.path & "\data\variables.ini"
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = GetVar(filename, "Variables", "Variable" & CStr(i) & "Name")
    Next
End Sub
