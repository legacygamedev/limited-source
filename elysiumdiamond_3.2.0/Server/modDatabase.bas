Attribute VB_Name = "modDatabase"

' Copyright (c) 2007-2008 Elysium Source. All rights reserved.
' This code is licensed under the Elysium General License.
Option Explicit

Public Declare Function WritePrivateProfileString& Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal FileName$)
Public Declare Function GetPrivateProfileString& Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal AppName$, ByVal KeyName$, ByVal keydefault$, ByVal ReturnedString$, ByVal RSSize&, ByVal FileName$)

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Public Const ADMIN_LOG As String = "logs\admin.txt"
Public Const PLAYER_LOG As String = "logs\player.txt"

Function AccountExist(ByVal Name As String) As Boolean
    Dim FileName As String

    FileName = "Accounts\" & Trim$(Name) & ".ini"

    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If

End Function

Sub AddAccount(ByVal Index As Long, _
   ByVal Name As String, _
   ByVal Password As String)
    Dim i As Long

    Player(Index).Login = Name
    Player(Index).Password = Password

    For i = 1 To MAX_CHARS
        Call ClearChar(Index, i)
    Next

    Call SavePlayer(Index)
End Sub

Sub AddChar(ByVal Index As Long, _
   ByVal Name As String, _
   ByVal Sex As Byte, _
   ByVal ClassNum As Byte, _
   ByVal CharNum As Long)
    Dim f As Long

    If Trim$(Player(Index).Char(CharNum).Name) = vbNullString Then
        Player(Index).CharNum = CharNum
        Player(Index).Char(CharNum).Name = Name
        Player(Index).Char(CharNum).Sex = Sex
        Player(Index).Char(CharNum).Class = ClassNum

        If Player(Index).Char(CharNum).Sex = SEX_MALE Then
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).MaleSprite
        Else
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).FemaleSprite
        End If

        Player(Index).Char(CharNum).Level = 1
        Player(Index).Char(CharNum).STR = Class(ClassNum).STR
        Player(Index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(Index).Char(CharNum).Speed = Class(ClassNum).Speed
        Player(Index).Char(CharNum).Magi = Class(ClassNum).Magi

        If Class(ClassNum).Map <= 0 Then Class(ClassNum).Map = 1
        If Class(ClassNum).X < 0 Or Class(ClassNum).X > MAX_MAPX Then Class(ClassNum).X = Int(Class(ClassNum).X / 2)
        If Class(ClassNum).Y < 0 Or Class(ClassNum).Y > MAX_MAPY Then Class(ClassNum).Y = Int(Class(ClassNum).Y / 2)
        Player(Index).Char(CharNum).Map = Class(ClassNum).Map
        Player(Index).Char(CharNum).X = Class(ClassNum).X
        Player(Index).Char(CharNum).Y = Class(ClassNum).Y
        Player(Index).Char(CharNum).HP = GetPlayerMaxHP(Index)
        Player(Index).Char(CharNum).MP = GetPlayerMaxMP(Index)
        Player(Index).Char(CharNum).SP = GetPlayerMaxSP(Index)

        ' Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
        Print #f, Name
        Close #f
        Call SavePlayer(Index)
        Exit Sub
    End If

End Sub

Sub AddLog(ByVal text As String, _
   ByVal FN As String)
    Dim FileName As String
    Dim f As Long

    If ServerLog = True Then
        FileName = App.Path & "\" & FN

        If Not FileExist(FN) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If

        f = FreeFile
        Open FileName For Append As #f
        Print #f, Time & ": " & text
        Close #f
    End If

End Sub

Sub BanByServer(ByVal BanPlayerIndex As Long, _
   ByVal Reason As String)
    Dim FileName, ip As String
    Dim f As Long, i As Long

    FileName = App.Path & "\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If

    ' Cut off last portion of ip
    ip = GetPlayerIP(BanPlayerIndex)

    For i = Len(ip) To 1 Step -1

        If Mid$(ip, i, 1) = "." Then
            Exit For
        End If

    Next

    ip = Mid$(ip, 1, i)
    f = FreeFile
    Open FileName For Append As #f
    Print #f, ip & "," & "Server"
    Close #f

    If Trim$(Reason) <> vbNullString Then
        Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by the server! Reason(" & Reason & ")", White)
        Call AlertMsg(BanPlayerIndex, "You have been banned by the server!  Reason(" & Reason & ")")
        Call AddLog("The server has banned " & GetPlayerName(BanPlayerIndex) & ".  Reason(" & Reason & ")", ADMIN_LOG)
    Else
        Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by the server!", White)
        Call AlertMsg(BanPlayerIndex, "You have been banned by the server!")
        Call AddLog("The server has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    End If

End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, _
   ByVal BannedByIndex As Long)
    Dim FileName, ip As String
    Dim f As Long, i As Long

    FileName = App.Path & "\banlist.txt"

    ' Make sure the file exists
    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If

    ' Cut off last portion of ip
    ip = GetPlayerIP(BanPlayerIndex)

    For i = Len(ip) To 1 Step -1

        If Mid$(ip, i, 1) = "." Then
            Exit For
        End If

    Next

    ip = Mid$(ip, 1, i)
    f = FreeFile
    Open FileName For Append As #f
    Print #f, ip & "," & GetPlayerName(BannedByIndex)
    Close #f
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean

    If Trim$(Player(Index).Char(CharNum).Name) <> vbNullString Then
        CharExist = True
    Else
        CharExist = False
    End If

End Function

Sub CheckArrows()

    If Not FileExist("Arrows.ini") Then
        Dim i As Long

        For i = 1 To MAX_ARROWS
            Call SetStatus("Saving arrows... " & Int((i / MAX_ARROWS) * 100) & "%")

            DoEvents
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowName", vbNullString)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowPic", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowRange", 0)
        Next

    End If

End Sub

Sub CheckClasses()

    If Not FileExist("Classes\info.ini") Then
        Call SaveClasses
    End If

End Sub

Sub CheckEmos()

    If Not FileExist("emoticons.ini") Then
        Dim i As Long

        For i = 0 To MAX_EMOTICONS
            Call SetStatus("Saving emoticons... " & Int((i / MAX_EMOTICONS) * 100) & "%")

            DoEvents
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "Emoticon" & i, 0)
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "EmoticonT" & i, 0)
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "EmoticonS" & i, 0)
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "EmoticonC" & i, vbNullString)
        Next

    End If

End Sub

Sub CheckExps()

    If Not FileExist("experience.ini") Then
        Dim i As Long

        For i = 1 To MAX_LEVEL
            Call SetStatus("Saving exp... " & Int((i / MAX_LEVEL) * 100) & "%")

            DoEvents
            Call PutVar(App.Path & "\experience.ini", "EXPERIENCE", "Exp" & i, i * 1500)
        Next

    End If

End Sub

Sub CheckItems()
    Call SaveItems
End Sub

Sub CheckMaps()
    Dim FileName As String

    Call ClearMaps
    Dim i As Long

    For i = 1 To MAX_MAPS
        FileName = "maps\map" & i & ".dat"

        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            Call SetStatus("Saving maps... " & Int((i / MAX_MAPS) * 100) & "%")

            DoEvents
            Call SaveMap(i)
        End If

    Next

End Sub

Sub CheckNpcs()
    Call SaveNpcs
End Sub

Sub CheckShops()
    Call SaveShops
End Sub

Sub CheckSpeech()
    Call SaveSpeeches
End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub ClearArrows()
    Dim i As Long

    For i = 1 To MAX_ARROWS
        Arrows(i).Name = vbNullString
        Arrows(i).Pic = 0
        Arrows(i).Range = 0
    Next

End Sub

Sub ClearEmos()
    Dim i As Long

    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Type = 0
        Emoticons(i).Pic = 0
        Emoticons(i).sound = vbNullString
        Emoticons(i).Command = vbNullString
    Next

End Sub

Sub ClearExps()
    Dim i As Long

    For i = 1 To MAX_LEVEL
        Experience(i) = 0
    Next

End Sub

Sub ClearParties()
    Dim i, o As Long

    For i = 1 To MAX_PARTIES
        For o = 1 To MAX_PARTY_MEMBERS
            Party(i).Member(o) = 0
        Next
    Next

End Sub

Sub DelChar(ByVal Index As Long, _
   ByVal CharNum As Long)
    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long, f2 As Long
    Dim s As String

    Call FileCopy(App.Path & "\accounts\charlist.txt", App.Path & "\accounts\chartemp.txt")

    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\accounts\chartemp.txt")
End Sub

Public Sub DelVar(sFileName As String, _
   sSection As String, _
   sKey As String)

    If Len(Trim$(sKey)) <> 0 Then
        WritePrivateProfileString sSection, sKey, vbNullString, sFileName
    Else
        WritePrivateProfileString sSection, sKey, vbNullString, sFileName
    End If

End Sub

Public Function ExistVar(File As String, Header As String, Var As String) As Boolean

    ExistVar = (GetVar(File, Header, Var) <> "")

End Function

Function FileExist(ByVal FileName As String) As Boolean

    If Dir$(App.Path & "\" & FileName) = vbNullString Then
        FileExist = False
    Else
        FileExist = True
    End If

End Function

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim s As String

    FindChar = False
    f = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If

    Loop

    Close #f
End Function

Sub LoadArrows()
    Dim FileName As String
    Dim i As Long

    Call CheckArrows
    FileName = App.Path & "\Arrows.ini"

    For i = 1 To MAX_ARROWS
        Call SetStatus("Loading Arrows... " & Int((i / MAX_ARROWS) * 100) & "%")
        Arrows(i).Name = GetVar(FileName, "Arrow" & i, "ArrowName")
        Arrows(i).Pic = GetVar(FileName, "Arrow" & i, "ArrowPic")
        Arrows(i).Range = GetVar(FileName, "Arrow" & i, "ArrowRange")

        DoEvents
    Next

End Sub

Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Sub LoadClasses()
    Dim FileName As String
    Dim i As Long

    Call CheckClasses
    FileName = App.Path & "\Classes\info.ini"
    Max_Classes = Val(GetVar(FileName, "INFO", "MaxClasses"))
    ReDim Class(1 To Max_Classes) As ClassRec
    Call ClearClasses

    For i = 1 To Max_Classes
        Call SetStatus("Loading classes... " & Int((i / Max_Classes) * 100) & "%")
        FileName = App.Path & "\Classes\Class" & i & ".ini"
        Class(i).Name = GetVar(FileName, "CLASS", "Name")
        Class(i).MaleSprite = GetVar(FileName, "CLASS", "MaleSprite")
        Class(i).FemaleSprite = GetVar(FileName, "CLASS", "FemaleSprite")
        Class(i).STR = Val(GetVar(FileName, "CLASS", "str"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS", "DEF"))
        Class(i).Speed = Val(GetVar(FileName, "CLASS", "SPEED"))
        Class(i).Magi = Val(GetVar(FileName, "CLASS", "MAGI"))
        Class(i).Map = Val(GetVar(FileName, "CLASS", "MAP"))
        Class(i).X = Val(GetVar(FileName, "CLASS", "X"))
        Class(i).Y = Val(GetVar(FileName, "CLASS", "Y"))
        Class(i).Locked = Val(GetVar(FileName, "CLASS", "Locked"))

        DoEvents
    Next

End Sub

Sub LoadEmos()
    Dim FileName As String
    Dim i As Long

    Call CheckEmos
    FileName = App.Path & "\emoticons.ini"

    For i = 0 To MAX_EMOTICONS
        Call SetStatus("Loading emoticons... " & Int((i / MAX_EMOTICONS) * 100) & "%")
        Emoticons(i).Type = Val(GetVar(FileName, "EMOTICONS", "EmoticonT" & i))
        Emoticons(i).Pic = Val(GetVar(FileName, "EMOTICONS", "Emoticon" & i))
        Emoticons(i).sound = GetVar(FileName, "EMOTICONS", "EmoticonS" & i)
        Emoticons(i).Command = GetVar(FileName, "EMOTICONS", "EmoticonC" & i)

        DoEvents
    Next

End Sub

Sub LoadExps()
    Dim FileName As String
    Dim i As Long

    Call CheckExps
    FileName = App.Path & "\experience.ini"

    For i = 1 To MAX_LEVEL
        Call SetStatus("Loading exp... " & Int((i / MAX_LEVEL) * 100) & "%")
        Experience(i) = GetVar(FileName, "EXPERIENCE", "Exp" & i)

        DoEvents
    Next

End Sub

Sub LoadItems()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckItems

    For i = 1 To MAX_ITEMS
        Call SetStatus("Loading items... " & Int((i / MAX_ITEMS) * 100) & "%")
        FileName = App.Path & "\Items\Item" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Item(i)
        Close #f

        DoEvents
    Next

End Sub

Sub LoadMaps()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckMaps

    For i = 1 To MAX_MAPS
        Call SetStatus("Loading maps... " & Int((i / MAX_MAPS) * 100) & "%")
        FileName = App.Path & "\maps\map" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Map(i)
        Close #f

        DoEvents
    Next

End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckNpcs

    For i = 1 To MAX_NPCS
        Call SetStatus("Loading npcs... " & Int((i / MAX_NPCS) * 100) & "%")
        FileName = App.Path & "\npcs\npc" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Npc(i)
        Close #f

        DoEvents
    Next

End Sub

Sub LoadPlayer(ByVal Index As Long, _
   ByVal Name As String)
    Dim FileName As String
    Dim i As Long
    Dim N As Long

    Call ClearPlayer(Index)
    FileName = App.Path & "\accounts\" & Trim$(Name) & ".ini"
    Player(Index).Login = GetVar(FileName, "GENERAL", "Login")
    Player(Index).Password = GetVar(FileName, "GENERAL", "Password")
    Player(Index).Pet.Alive = NO

    For i = 1 To MAX_CHARS

        ' General
        Player(Index).Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
        Player(Index).Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
        Player(Index).Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))

        If Player(Index).Char(i).Class = 0 Then Player(Index).Char(i).Class = 1
        Player(Index).Char(i).Sprite = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
        Player(Index).Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
        Player(Index).Char(i).Exp = Val(GetVar(FileName, "CHAR" & i, "Exp"))
        Player(Index).Char(i).Access = Val(GetVar(FileName, "CHAR" & i, "Access"))
        Player(Index).Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
        Player(Index).Char(i).Guild = GetVar(FileName, "CHAR" & i, "Guild")
        Player(Index).Char(i).Guildaccess = Val(GetVar(FileName, "CHAR" & i, "Guildaccess"))

        ' Vitals
        Player(Index).Char(i).HP = Val(GetVar(FileName, "CHAR" & i, "HP"))
        Player(Index).Char(i).MP = Val(GetVar(FileName, "CHAR" & i, "MP"))
        Player(Index).Char(i).SP = Val(GetVar(FileName, "CHAR" & i, "SP"))

        ' Stats
        Player(Index).Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "str"))
        Player(Index).Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
        Player(Index).Char(i).Speed = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        Player(Index).Char(i).Magi = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
        Player(Index).Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))

        ' Worn equipment
        Player(Index).Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
        Player(Index).Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
        Player(Index).Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
        Player(Index).Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))

        ' Position
        Player(Index).Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
        Player(Index).Char(i).X = Val(GetVar(FileName, "CHAR" & i, "X"))
        Player(Index).Char(i).Y = Val(GetVar(FileName, "CHAR" & i, "Y"))
        Player(Index).Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))

        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).X = START_X
            Player(Index).Char(i).Y = START_Y
        End If

        ' Inventory
        For N = 1 To MAX_INV
            Player(Index).Char(i).Inv(N).num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & N))
            Player(Index).Char(i).Inv(N).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & N))
            Player(Index).Char(i).Inv(N).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & N))
        Next

        ' Spells
        For N = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(i).Spell(N) = Val(GetVar(FileName, "CHAR" & i, "Spell" & N))
        Next

        If Val(GetVar(FileName, "CHAR" & i, "HasPet")) = 1 Then
            Player(Index).Pet.Sprite = Val(GetVar(FileName, "CHAR" & i, "Pet"))
            Player(Index).Pet.Alive = YES
            Player(Index).Pet.Dir = DIR_UP
            Player(Index).Pet.Map = Player(Index).Char(i).Map
            Player(Index).Pet.X = Player(Index).Char(i).X + Int((Rnd * 3) - 1)

            If Player(Index).Pet.X < 0 Or Player(Index).Pet.X > MAX_MAPX Then Player(Index).Pet.X = GetPlayerX(Index)
            Player(Index).Pet.Y = Player(Index).Char(i).Y + Int((Rnd * 3) - 1)

            If Player(Index).Pet.Y < 0 Or Player(Index).Pet.Y > MAX_MAPY Then Player(Index).Pet.Y = GetPlayerY(Index)
            Player(Index).Pet.MapToGo = 0
            Player(Index).Pet.XToGo = -1
            Player(Index).Pet.YToGo = -1
            Player(Index).Pet.Level = Val(GetVar(FileName, "CHAR" & i, "PetLevel"))
            Player(Index).Pet.HP = Player(Index).Pet.Level * 5 '???
        End If

        For N = 1 To MAX_FRIENDS
            Player(Index).Char(i).Friends(N) = GetVar(FileName, "CHAR" & i, "Friend" & N)
        Next
    Next

End Sub

Sub LoadShops()
    Dim FileName As String
    Dim i As Long, f As Long

    Call CheckShops

    For i = 1 To MAX_SHOPS
        Call SetStatus("Loading shops... " & Int((i / MAX_SHOPS) * 100) & "%")
        FileName = App.Path & "\shops\shop" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Shop(i)
        Close #f

        DoEvents
    Next

End Sub

Sub LoadSpeeches()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckSpeech

    For i = 1 To MAX_SPEECH
        Call SetStatus("Loading speech... " & Int((i / MAX_SPEECH) * 100) & "%")
        FileName = App.Path & "\speech\speech" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Speech(i)
        Close #f

        DoEvents
    Next

End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim i As Long
    Dim f As Long

    Call CheckSpells

    For i = 1 To MAX_SPELLS
        Call SetStatus("Loading spells... " & Int((i / MAX_SPELLS) * 100) & "%")
        FileName = App.Path & "\spells\spells" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Spell(i)
        Close #f

        DoEvents
    Next

End Sub

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim FileName As String
    Dim RightPassword As String

    PasswordOK = False

    If AccountExist(Name) Then
        FileName = App.Path & "\accounts\" & Trim$(Name) & ".ini"
        RightPassword = GetVar(FileName, "GENERAL", "Password")

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub PutVar(File As String, _
   Header As String, _
   Var As String, _
   Value As String)

    If Trim$(Value) = "0" Or Trim$(Value) = vbNullString Then
        If ExistVar(File, Header, Var) Then
            Call DelVar(File, Header, Var)
        End If

    Else
        Call WritePrivateProfileString(Header, Var, Value, File)
    End If

End Sub

Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If

    Next

End Sub

Sub SaveArrow(ByVal ArrowNum As Long)
    Dim FileName As String

    FileName = App.Path & "\Arrows.ini"
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowName", Trim$(Arrows(ArrowNum).Name))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range))
End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim i As Long

    FileName = App.Path & "\Classes\info.ini"

    If Not FileExist("Classes\info.ini") Then
        Call SpecialPutVar(FileName, "INFO", "MaxClasses", 3)
        Max_Classes = 3
    End If

    For i = 1 To Max_Classes
        Call SetStatus("Saving classes... " & Int((i / Max_Classes) * 100) & "%")

        DoEvents
        FileName = App.Path & "\Classes\Class" & i & ".ini"

        If Not FileExist("Classes\Class" & i & ".ini") Then
            Call PutVar(FileName, "CLASS", "Name", Trim$(Class(i).Name))
            Call PutVar(FileName, "CLASS", "MaleSprite", STR(Class(i).MaleSprite))
            Call PutVar(FileName, "CLASS", "FemaleSprite", STR(Class(i).FemaleSprite))
            Call PutVar(FileName, "CLASS", "str", STR(Class(i).STR))
            Call PutVar(FileName, "CLASS", "DEF", STR(Class(i).DEF))
            Call PutVar(FileName, "CLASS", "SPEED", STR(Class(i).Speed))
            Call PutVar(FileName, "CLASS", "MAGI", STR(Class(i).Magi))
            Call PutVar(FileName, "CLASS", "MAP", STR(Class(i).Map))
            Call PutVar(FileName, "CLASS", "X", STR(Class(i).X))
            Call PutVar(FileName, "CLASS", "Y", STR(Class(i).Y))
            Call PutVar(FileName, "CLASS", "Locked", STR(Class(i).Locked))
        End If

    Next

End Sub

Sub SaveEmoticon(ByVal EmoNum As Long)
    Dim FileName As String

    FileName = App.Path & "\emoticons.ini"
    Call PutVar(FileName, "EMOTICONS", "EmoticonT" & EmoNum, STR(Emoticons(EmoNum).Type))
    Call PutVar(FileName, "EMOTICONS", "EmoticonC" & EmoNum, Trim$(Emoticons(EmoNum).Command))
    Call PutVar(FileName, "EMOTICONS", "Emoticon" & EmoNum, Val(Emoticons(EmoNum).Pic))
    Call PutVar(FileName, "EMOTICONS", "EmoticonS" & EmoNum, Emoticons(EmoNum).sound)
End Sub

Sub SaveItem(ByVal ItemNum As Long)
    Dim FileName As String
    Dim f  As Long

    FileName = App.Path & "\items\item" & ItemNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub SaveItems()
    Dim i As Long

    Call SetStatus("Saving items... ")

    For i = 1 To MAX_ITEMS

        If Not FileExist("items\item" & i & ".dat") Then
            Call SetStatus("Saving items... " & Int((i / MAX_ITEMS) * 100) & "%")

            DoEvents
            Call SaveItem(i)
        End If

    Next

End Sub

Sub SaveLogs()
    Dim FileName As String
    Dim i As String, C As String

    If LCase$(Dir$(App.Path & "\logs", vbDirectory)) <> "logs" Then
        Call MkDir$(App.Path & "\Logs")
    End If

    'C = Time
    'C = Replace(C, ":", ".", 1)
    'C = Replace(C, ":", ".", 1)
    'C = Replace(C, ":", ".", 1)
    'i = Date
    'i = Replace(i, "/", ".", 1)
    'i = Replace(i, "/", ".", 1)
    'i = Replace(i, "/", ".", 1)

    C = C & Hour(Time) & "." & Minute(Time) & "." & Second(Time)
   
    i = i & Year(Date) & "." & Month(Date) & "." & Day(Date)

    If LCase$(Dir$(App.Path & "\logs\" & i, vbDirectory)) <> i Then
        Call MkDir$(App.Path & "\Logs\" & i & "\")
    End If

    If LCase$(Dir$(App.Path & "\logs\" & i & "\" & C, vbDirectory)) <> C Then
        Call MkDir$(App.Path & "\Logs\" & i & "\" & C & "\")
    End If

    'FileName = App.Path & "\Logs\" & i & "\" & C & "\Main.txt"
    'Open FileName For Output As #1
    'Print #1, frmServer.txtText(0).text
    'Close #1
    'FileName = App.Path & "\Logs\" & i & "\" & C & "\Broadcast.txt"
    'Open FileName For Output As #1
    'Print #1, frmServer.txtText(1).text
    'Close #1
    'FileName = App.Path & "\Logs\" & i & "\" & C & "\Global.txt"
    'Open FileName For Output As #1
    'Print #1, frmServer.txtText(2).text
    'Close #1
    'FileName = App.Path & "\Logs\" & i & "\" & C & "\Map.txt"
    'Open FileName For Output As #1
    'Print #1, frmServer.txtText(3).text
    'Close #1
    'FileName = App.Path & "\Logs\" & i & "\" & C & "\Private.txt"
    'Open FileName For Output As #1
    'Print #1, frmServer.txtText(4).text
    'Close #1
    'FileName = App.Path & "\Logs\" & i & "\" & C & "\Admin.txt"
    'Open FileName For Output As #1
    'Print #1, frmServer.txtText(5).text
    'Close #1
    'FileName = App.Path & "\Logs\" & i & "\" & C & "\Emote.txt"
    'Open FileName For Output As #1
    'Print #1, frmServer.txtText(6).text
    'Close #1
End Sub

Sub SaveMap(ByVal MapNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\maps\map" & MapNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\npcs\npc" & NpcNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Npc(NpcNum)
    Close #f
End Sub

Sub SaveNpcs()
    Dim i As Long

    Call SetStatus("Saving npcs... ")

    For i = 1 To MAX_NPCS

        If Not FileExist("npcs\npc" & i & ".dat") Then
            Call SetStatus("Saving npcs... " & Int((i / MAX_NPCS) * 100) & "%")

            DoEvents
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim FileName As String
    Dim i As Long
    Dim N As Long

    FileName = App.Path & "\accounts\" & Trim$(Player(Index).Login) & ".ini"
    Call PutVar(FileName, "GENERAL", "Login", Trim$(Player(Index).Login))
    Call PutVar(FileName, "GENERAL", "Password", Trim$(Player(Index).Password))

    For i = 1 To MAX_CHARS

        ' General
        Call PutVar(FileName, "CHAR" & i, "Name", Trim$(Player(Index).Char(i).Name))
        Call PutVar(FileName, "CHAR" & i, "Class", STR(Player(Index).Char(i).Class))
        Call PutVar(FileName, "CHAR" & i, "Sex", STR(Player(Index).Char(i).Sex))
        Call PutVar(FileName, "CHAR" & i, "Sprite", STR(Player(Index).Char(i).Sprite))
        Call PutVar(FileName, "CHAR" & i, "Level", STR(Player(Index).Char(i).Level))
        Call PutVar(FileName, "CHAR" & i, "Exp", STR(Player(Index).Char(i).Exp))
        Call PutVar(FileName, "CHAR" & i, "Access", STR(Player(Index).Char(i).Access))
        Call PutVar(FileName, "CHAR" & i, "PK", STR(Player(Index).Char(i).PK))
        Call PutVar(FileName, "CHAR" & i, "Guild", Trim$(Player(Index).Char(i).Guild))
        Call PutVar(FileName, "CHAR" & i, "Guildaccess", STR(Player(Index).Char(i).Guildaccess))

        ' Vitals
        Call PutVar(FileName, "CHAR" & i, "HP", STR(Player(Index).Char(i).HP))
        Call PutVar(FileName, "CHAR" & i, "MP", STR(Player(Index).Char(i).MP))
        Call PutVar(FileName, "CHAR" & i, "SP", STR(Player(Index).Char(i).SP))

        ' Stats
        Call PutVar(FileName, "CHAR" & i, "str", STR(Player(Index).Char(i).STR))
        Call PutVar(FileName, "CHAR" & i, "DEF", STR(Player(Index).Char(i).DEF))
        Call PutVar(FileName, "CHAR" & i, "SPEED", STR(Player(Index).Char(i).Speed))
        Call PutVar(FileName, "CHAR" & i, "MAGI", STR(Player(Index).Char(i).Magi))
        Call PutVar(FileName, "CHAR" & i, "POINTS", STR(Player(Index).Char(i).POINTS))

        ' Worn equipment
        Call PutVar(FileName, "CHAR" & i, "ArmorSlot", STR(Player(Index).Char(i).ArmorSlot))
        Call PutVar(FileName, "CHAR" & i, "WeaponSlot", STR(Player(Index).Char(i).WeaponSlot))
        Call PutVar(FileName, "CHAR" & i, "HelmetSlot", STR(Player(Index).Char(i).HelmetSlot))
        Call PutVar(FileName, "CHAR" & i, "ShieldSlot", STR(Player(Index).Char(i).ShieldSlot))

        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).X = START_X
            Player(Index).Char(i).Y = START_Y
        End If

        ' Position
        Call PutVar(FileName, "CHAR" & i, "Map", STR(Player(Index).Char(i).Map))
        Call PutVar(FileName, "CHAR" & i, "X", STR(Player(Index).Char(i).X))
        Call PutVar(FileName, "CHAR" & i, "Y", STR(Player(Index).Char(i).Y))
        Call PutVar(FileName, "CHAR" & i, "Dir", STR(Player(Index).Char(i).Dir))

        ' Inventory
        For N = 1 To MAX_INV
            Call PutVar(FileName, "CHAR" & i, "InvItemNum" & N, STR(Player(Index).Char(i).Inv(N).num))
            Call PutVar(FileName, "CHAR" & i, "InvItemVal" & N, STR(Player(Index).Char(i).Inv(N).Value))
            Call PutVar(FileName, "CHAR" & i, "InvItemDur" & N, STR(Player(Index).Char(i).Inv(N).Dur))
        Next

        ' Spells
        For N = 1 To MAX_PLAYER_SPELLS
            Call PutVar(FileName, "CHAR" & i, "Spell" & N, STR(Player(Index).Char(i).Spell(N)))
        Next

        ' Pet
        If i = Player(Index).CharNum Then
            If Player(Index).Pet.Alive = YES Then
                Call PutVar(FileName, "CHAR" & i, "HasPet", 1)
                Call PutVar(FileName, "CHAR" & i, "Pet", STR(Player(Index).Pet.Sprite))
                Call PutVar(FileName, "CHAR" & i, "PetLevel", STR(Player(Index).Pet.Level))
            Else
                Call PutVar(FileName, "CHAR" & i, "HasPet", 0)
                Call DelVar(FileName, "CHAR" & i, "Pet") ' Saving space
                Call DelVar(FileName, "CHAR" & i, "PetLevel")
            End If

        Else
            Call PutVar(FileName, "CHAR" & i, "HasPet", 0)
            Call DelVar(FileName, "CHAR" & i, "Pet") ' Saving space
            Call DelVar(FileName, "CHAR" & i, "PetLevel")
        End If

        ' Friend list
        For N = 1 To MAX_FRIENDS
            Call PutVar(FileName, "CHAR" & i, "Friend" & N, Player(Index).Char(i).Friends(N))
        Next
    Next

End Sub

Sub SaveShop(ByVal ShopNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\shops\shop" & ShopNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub SaveShops()
    Dim i As Long

    Call SetStatus("Saving shops... ")

    For i = 1 To MAX_SHOPS

        If Not FileExist("shops\shop" & i & ".dat") Then
            Call SetStatus("Saving shops... " & Int((i / MAX_SHOPS) * 100) & "%")

            DoEvents
            Call SaveShop(i)
        End If

    Next

End Sub

Sub SaveSpeech(ByVal Index As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\speech\speech" & Index & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Speech(Index)
    Close #f
End Sub

Sub SaveSpeeches()
    Dim i As Long

    Call SetStatus("Saving speech... ")

    For i = 1 To MAX_SPEECH

        If Not FileExist("speech\speech" & i & ".dat") Then
            Call SetStatus("Saving speech... " & Int((i / MAX_SPEECH) * 100) & "%")

            DoEvents
            Call SaveSpeech(i)
        End If

    Next

End Sub

Sub SaveSpell(ByVal SpellNum As Long)
    Dim FileName As String
    Dim f As Long

    FileName = App.Path & "\spells\spells" & SpellNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
    Dim i As Long

    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS

        If Not FileExist("spells\spells" & i & ".dat") Then
            Call SetStatus("Saving spells... " & Int((i / MAX_SPELLS) * 100) & "%")

            DoEvents
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub SpecialPutVar(File As String, _
   Header As String, _
   Var As String, _
   Value As String)

    ' Same as the one below except it keeps all 0 and blank values (used for config)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Private Function Replace(strWord, _
   strFind, _
   strReplace, _
   charAmount) As String
    Dim a  As Integer

    a = InStr(1, UCase$(strWord), UCase$(strFind))
    On Error Resume Next
    strWord = Mid$(strWord, 1, a - 1) & strReplace & Right$(strWord, Len(strWord) - a - charAmount + 1)
    Replace = strWord
End Function
