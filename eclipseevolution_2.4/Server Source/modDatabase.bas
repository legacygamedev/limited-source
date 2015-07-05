Attribute VB_Name = "modDatabase"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Public temp As Integer

Private LastPercent As Long

Public Const ADMIN_LOG = "logs\admin.txt"
Public Const PLAYER_LOG = "logs\player.txt"

Public Sub UpdatePercentage(Percent As Integer, Message As String)

    '//!! Displays the loading percentages in divisibles of 10 only - much faster since theres lots less GUI refreshing
    If Abs(LastPercent - Percent) > 5 Then LastPercent = 0
    If Percent > LastPercent Then
        If Percent Mod 10 = 0 Then
            LastPercent = Percent + 1
            Call SetStatus(Message & " " & Percent & "%")
            DoEvents
        End If
    End If

End Sub

Function AccountExist(ByVal Name As String) As Boolean

  Dim FileName As String

    FileName = "accounts\" & Trim(Name) & ".ini"

    If FileExist(FileName) Then
        AccountExist = True
        Exit Function
     Else
        AccountExist = False
    End If

    FileName = "accounts\" & Trim(Name) & "_info.ini"

    If FileExist(FileName) Then
        AccountExist = True
        Exit Function
     Else
        AccountExist = False
    End If

End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String, ByVal Email As String)

  Dim i As Long

    Player(index).Login = Name
    Player(index).Password = Password
    Player(index).Email = Email

    For i = 1 To MAX_CHARS
        Call ClearChar(index, i)
    Next i

    Call SavePlayer(index)

    If Val(GetVar(App.Path & "\Data.ini", "CONFIG", "verified")) = 1 Then
        Call PutVar(App.Path & "\accounts\" & Trim(Player(index).Login) & "_info.ini", "ACCESS", "verified", 0)
    End If

End Sub

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long, ByVal headc As Long, ByVal bodyc As Long, ByVal logc As Long)

  Dim f As Long

    If Trim(Player(index).Char(CharNum).Name) = "" Then
        Player(index).CharNum = CharNum

        Player(index).Char(CharNum).Name = Name
        Player(index).Char(CharNum).Sex = Sex
        Player(index).Char(CharNum).Class = ClassNum

        If Player(index).Char(CharNum).Sex = SEX_MALE Then
            Player(index).Char(CharNum).Sprite = Class(ClassNum).MaleSprite
         Else
            Player(index).Char(CharNum).Sprite = Class(ClassNum).FemaleSprite
        End If

        Player(index).Char(CharNum).Level = 1

        Player(index).Char(CharNum).STR = Class(ClassNum).STR
        Player(index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(index).Char(CharNum).Speed = Class(ClassNum).Speed
        Player(index).Char(CharNum).Magi = Class(ClassNum).Magi

        If Class(ClassNum).Map <= 0 Then Class(ClassNum).Map = 1
        If Class(ClassNum).X < 0 Or Class(ClassNum).X > MAX_MAPX Then Class(ClassNum).X = Int(Class(ClassNum).X / 2)
        If Class(ClassNum).Y < 0 Or Class(ClassNum).Y > MAX_MAPY Then Class(ClassNum).Y = Int(Class(ClassNum).Y / 2)
        Player(index).Char(CharNum).Map = Class(ClassNum).Map
        Player(index).Char(CharNum).X = Class(ClassNum).X
        Player(index).Char(CharNum).Y = Class(ClassNum).Y

        Player(index).Char(CharNum).HP = GetPlayerMaxHP(index)
        Player(index).Char(CharNum).MP = GetPlayerMaxMP(index)
        Player(index).Char(CharNum).SP = GetPlayerMaxSP(index)

        Player(index).Char(CharNum).head = headc
        Player(index).Char(CharNum).body = bodyc
        Player(index).Char(CharNum).leg = logc

        Player(index).Char(CharNum).Paperdoll = 1

        ' Append name to file
        f = FreeFile
        Open App.Path & "\accounts\charlist.txt" For Append As #f
        Print #f, Name
        Close #f

        Call SavePlayer(index)

        Exit Sub
    End If

End Sub

Sub AddLog(ByVal text As String, ByVal FN As String)

    On Error Resume Next
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

Sub BanByServer(ByVal BanPlayerIndex As Long, ByVal Reason As String)

  Dim FileName
  Dim IP As String
  Dim f As Long
  Dim i As Long

    FileName = App.Path & "\banlist.txt"

    ' Make sure the file exists

    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid(IP, i, 1) = "." Then
            Exit For
        End If

    Next i
    IP = Mid(IP, 1, i)

    f = FreeFile
    Open FileName For Append As #f
    Print #f, IP & "," & "Server"
    Close #f

    If Trim(Reason) <> "" Then
        Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by the server! Reason(" & Reason & ")", White)
        Call AlertMsg(BanPlayerIndex, "You have been banned by the server!  Reason(" & Reason & ")")
        Call AddLog("The server has banned " & GetPlayerName(BanPlayerIndex) & ".  Reason(" & Reason & ")", ADMIN_LOG)
     Else
        Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by the server!", White)
        Call AlertMsg(BanPlayerIndex, "You have been banned by the server!")
        Call AddLog("The server has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    End If

End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)

  Dim FileName
  Dim IP As String
  Dim f As Long
  Dim i As Long

    FileName = App.Path & "\banlist.txt"

    ' Make sure the file exists

    If Not FileExist("banlist.txt") Then
        f = FreeFile
        Open FileName For Output As #f
        Close #f
    End If

    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)

    For i = Len(IP) To 1 Step -1

        If Mid(IP, i, 1) = "." Then
            Exit For
        End If

    Next i
    IP = Mid(IP, 1, i)

    f = FreeFile
    Open FileName For Append As #f
    Print #f, IP & "," & GetPlayerName(BannedByIndex)
    Close #f

    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")

End Sub

Function CharExist(ByVal index As Long, ByVal CharNum As Long) As Boolean

    If Trim(Player(index).Char(CharNum).Name) <> "" Then
        CharExist = True
     Else
        CharExist = False
    End If

End Function

Sub CheckArrows()

    If Not FileExist("Arrows.ini") Then
        Dim i As Long
        Dim Percent As Integer

        For i = 1 To MAX_ARROWS
            Percent = i / MAX_ARROWS * 100
            UpdatePercentage Percent, "Saving arrows..."
            
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowName", "")
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowPic", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowRange", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & i, "ArrowAmount", 0)
        Next i

    End If

End Sub

Sub CheckClasses()

    If Not FileExist("Classes\info.ini") Then
        Call SaveClasses
    End If

End Sub

Sub CheckClasses2()

    If Not FileExist("FirstClassAdvancement.ini") Then
        Call SaveClasses2
    End If

End Sub

Sub Checkclasses3()

    If Not FileExist("SecondClassAdvancement.ini") Then
        Call Saveclasses3
    End If

End Sub

Sub CheckElements()

    If Not FileExist("elements.ini") Then
        Dim i As Integer
        Dim Percent As Integer

        For i = 0 To MAX_ELEMENTS
            Percent = i / MAX_ELEMENTS * 100
            UpdatePercentage Percent, "Saving elements..."
            
            Call PutVar(App.Path & "\elements.ini", "ELEMENTS", "ElementName" & i, "")
            Call PutVar(App.Path & "\elements.ini", "ELEMENTS", "ElementStrong" & i, 0)
            Call PutVar(App.Path & "\elements.ini", "ELEMENTS", "ElementWeak" & i, 0)
        Next i

    End If

End Sub

Sub CheckEmos()

    If Not FileExist("emoticons.ini") Then
        Dim i As Integer

        For i = 0 To MAX_EMOTICONS
            temp = i / MAX_LEVEL * 100
            Call SetStatus("Saving emoticons... " & temp & "%")
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "Emoticon" & i, 0)
            Call PutVar(App.Path & "\emoticons.ini", "EMOTICONS", "EmoticonC" & i, "")
        Next i

    End If

End Sub

Sub CheckExps()

    If Not FileExist("experience.ini") Then
        Dim i As Integer

        For i = 1 To MAX_LEVEL
            temp = i / MAX_LEVEL * 100
            Call SetStatus("Saving exp... " & temp & "%")
            Call PutVar(App.Path & "\experience.ini", "EXPERIENCE", "Exp" & i, i * 1500)
        Next i

    End If

End Sub

Sub CheckItems()

    Call SaveItems

End Sub

Sub CheckMaps()

  Dim i As Integer
  Dim Percent As Integer

    Call ClearMaps

    For i = 1 To MAX_MAPS

        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist("maps\map" & i & ".dat") Then
            Percent = i / MAX_MAPS * 100
            UpdatePercentage Percent, "Saving maps..."

            'Save the map
            Call SaveMap(i)
            
        End If

    Next i

End Sub

Sub CheckNpcs()

    Call SaveNpcs

End Sub

Sub CheckQuests()

    Call SaveQuests

End Sub

Sub CheckShops()

    Call SaveShops

End Sub

Sub CheckSkills()

    Call SaveSkills

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
        Arrows(i).Amount = 0
    Next i

End Sub

Sub ClearEmos()

  Dim i As Integer

    For i = 0 To MAX_EMOTICONS
        Emoticons(i).Pic = 0
        Emoticons(i).Command = ""
    Next i

End Sub

Sub ClearExps()

  Dim i As Integer

    For i = 1 To MAX_LEVEL
        Experience(i) = 0
    Next i

End Sub

Sub DelChar(ByVal index As Long, ByVal CharNum As Long)

    Call DeleteName(Player(index).Char(CharNum).Name)
    Call ClearChar(index, CharNum)
    Call SavePlayer(index)

End Sub

Sub DeleteName(ByVal Name As String)

  Dim f1 As Long
  Dim f2 As Long
  Dim s As String

    Call FileCopy(App.Path & "\accounts\charlist.txt", App.Path & "\accounts\chartemp.txt")

    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim(LCase(s)) <> Trim(LCase(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2

    Call Kill(App.Path & "\accounts\chartemp.txt")

End Sub

Function FileExist(FileName As String) As Boolean

    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExist = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
ErrorHandler:
    ' if an error occurs, this function returns False

End Function

Function FindChar(ByVal Name As String) As Boolean

  Dim f As Long
  Dim s As String

    FindChar = False

    f = FreeFile
    Open App.Path & "\accounts\charlist.txt" For Input As #f

    Do While Not EOF(f)
        Input #f, s

        If Trim(LCase(s)) = Trim(LCase(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If

    Loop
    Close #f

End Function

Function GetVar(File As String, Header As String, Var As String) As String

  Dim sSpaces As String   ' Max string length

    On Error GoTo GetVar_Error

    sSpaces = Space(1000)

    Call GetPrivateProfileString(Header, Var, vbNullString, sSpaces, Len(sSpaces), File)

    GetVar = RTrim$(sSpaces)
    GetVar = left$(GetVar, Len(GetVar) - 1)

    On Error GoTo 0
    Exit Function

GetVar_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure GetVar of Module modDatabase"

End Function

Sub LoadArrows()

  Dim FileName As String
  Dim i As Long
  Dim Percent As Integer
  
    Call CheckArrows

    FileName = App.Path & "\Arrows.ini"

    For i = 1 To MAX_ARROWS
        Percent = i / MAX_ARROWS * 100
        UpdatePercentage Percent, "Loading Arrows..."
        
        Arrows(i).Name = GetVar(FileName, "Arrow" & i, "ArrowName")
        Arrows(i).Pic = GetVar(FileName, "Arrow" & i, "ArrowPic")
        Arrows(i).Range = GetVar(FileName, "Arrow" & i, "ArrowRange")
        Arrows(i).Amount = GetVar(FileName, "Arrow" & i, "ArrowAmount")

    Next i

End Sub

Sub LoadClasses()

  Dim FileName As String
  Dim i As Long
  Dim Percent As Integer

    On Error GoTo ClassErr
    Call CheckClasses

    FileName = App.Path & "\Classes\info.ini"

    MAX_CLASSES = Val(GetVar(FileName, "INFO", "MaxClasses"))

    ReDim Class(0 To MAX_CLASSES) As ClassRec

    Call ClearClasses

    For i = 0 To MAX_CLASSES
        Percent = i / MAX_CLASSES * 100
        UpdatePercentage Percent, "Loading classes..."
        
        FileName = App.Path & "\Classes\Class" & i & ".ini"

        ' Check if class exists

        If Not FileExist("\Classes\Class" & i & ".ini") Then
            Call PutVar(FileName, "CLASS", "Name", Trim(Class(i).Name))
            Call PutVar(FileName, "CLASS", "MaleSprite", STR(Class(i).MaleSprite))
            Call PutVar(FileName, "CLASS", "FemaleSprite", STR(Class(i).FemaleSprite))
            Call PutVar(FileName, "CLASS", "STR", STR(Class(i).STR))
            Call PutVar(FileName, "CLASS", "DEF", STR(Class(i).DEF))
            Call PutVar(FileName, "CLASS", "SPEED", STR(Class(i).Speed))
            Call PutVar(FileName, "CLASS", "MAGI", STR(Class(i).Magi))
            Call PutVar(FileName, "CLASS", "MAP", STR(Class(i).Map))
            Call PutVar(FileName, "CLASS", "X", STR(Class(i).X))
            Call PutVar(FileName, "CLASS", "Y", STR(Class(i).Y))
            Call PutVar(FileName, "CLASS", "Locked", STR(Class(i).locked))
            Call MsgBox("Class " & i & " not found, created.", vbInformation)
        End If

        Class(i).Name = GetVar(FileName, "CLASS", "Name")
        Class(i).MaleSprite = GetVar(FileName, "CLASS", "MaleSprite")
        Class(i).FemaleSprite = GetVar(FileName, "CLASS", "FemaleSprite")
        Class(i).STR = Val(GetVar(FileName, "CLASS", "STR"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS", "DEF"))
        Class(i).Speed = Val(GetVar(FileName, "CLASS", "SPEED"))
        Class(i).Magi = Val(GetVar(FileName, "CLASS", "MAGI"))
        Class(i).Map = Val(GetVar(FileName, "CLASS", "MAP"))
        Class(i).X = Val(GetVar(FileName, "CLASS", "X"))
        Class(i).Y = Val(GetVar(FileName, "CLASS", "Y"))
        Class(i).locked = Val(GetVar(FileName, "CLASS", "Locked"))

    Next i

    Exit Sub

ClassErr:
    Call MsgBox("Error loading class " & i & ". Check that all the variables in your class files exist!", vbCritical)
    Call DestroyServer
    End

End Sub

Sub LoadClasses2()

  Dim FileName As String
  Dim i As Long
  Dim Percent As Integer

    Call CheckClasses2

    FileName = App.Path & "\FirstClassAdvancement.ini"

    MAX_CLASSES = Val(GetVar(FileName, "INIT", "MaxClasses"))

    ReDim Class2(0 To MAX_CLASSES) As ClassRec

    Call ClearClasses2

    For i = 0 To MAX_CLASSES
        Percent = i / MAX_CLASSES * 100
        UpdatePercentage Percent, "Loading first class advancement..."
        
        Class2(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class2(i).AdvanceFrom = Val(GetVar(FileName, "CLASS" & i, "AdvanceFrom"))
        Class2(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class2(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        Class2(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class2(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class2(i).Speed = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class2(i).Magi = Val(GetVar(FileName, "CLASS" & i, "MAGI"))

    Next i

End Sub

Sub Loadclasses3()

  Dim FileName As String
  Dim i As Long
  Dim Percent As Integer

    Call Checkclasses3

    FileName = App.Path & "\SecondClassAdvancement.ini"

    MAX_CLASSES = Val(GetVar(FileName, "INIT", "MaxClasses"))

    ReDim Class3(0 To MAX_CLASSES) As ClassRec

    Call ClearClasses3

    For i = 0 To MAX_CLASSES
        Percent = i / MAX_CLASSES * 100
        UpdatePercentage Percent, "Loading second class advandcement..."

        Class3(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class3(i).AdvanceFrom = Val(GetVar(FileName, "CLASS" & i, "AdvanceFrom"))
        Class3(i).LevelReq = Val(GetVar(FileName, "CLASS" & i, "LevelReq"))
        Class3(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class3(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        Class3(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class3(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class3(i).Speed = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class3(i).Magi = Val(GetVar(FileName, "CLASS" & i, "MAGI"))

    Next i

End Sub

Sub LoadElements()

    On Error GoTo ElementErr
  Dim FileName As String
  Dim i As Integer
  Dim Percent As Integer

    Call CheckElements

    FileName = App.Path & "\elements.ini"

    For i = 0 To MAX_ELEMENTS
        Percent = i / MAX_ELEMENTS * 100
        UpdatePercentage Percent, "Loading elements..."
        Element(i).Name = GetVar(FileName, "ELEMENTS", "ElementName" & i)
        Element(i).Strong = Val(GetVar(FileName, "ELEMENTS", "ElementStrong" & i))
        Element(i).Weak = Val(GetVar(FileName, "ELEMENTS", "ElementWeak" & i))
    Next i

    Exit Sub

ElementErr:
    Call MsgBox("Error loading element " & i & ". Make sure all the variables in elements.ini are correct!", vbCritical)
    Call DestroyServer
    End

End Sub

Sub LoadEmos()

  Dim FileName As String
  Dim i As Integer
  Dim Percentage As Integer
  
    Call CheckEmos

    FileName = App.Path & "\emoticons.ini"

    For i = 0 To MAX_EMOTICONS
        Percentage = i / MAX_EMOTICONS * 100
        UpdatePercentage Percentage, "Loading emoticons..."

        Emoticons(i).Pic = GetVar(FileName, "EMOTICONS", "Emoticon" & i)
        Emoticons(i).Command = GetVar(FileName, "EMOTICONS", "EmoticonC" & i)
    Next i

End Sub

Sub LoadExps()

    On Error GoTo ExpErr
  Dim FileName As String
  Dim i As Integer
  Dim Percent As Integer

    Call CheckExps

    FileName = App.Path & "\experience.ini"

    For i = 1 To MAX_LEVEL
        Percent = i / MAX_LEVEL * 100
        UpdatePercentage Percent, "Loading exp..."

        Experience(i) = GetVar(FileName, "EXPERIENCE", "Exp" & i)
    Next i

    Exit Sub

ExpErr:
    Call MsgBox("Error loading EXP for level " & i & ". Make sure experience.ini has the correct variables! ERR: " & Err.number & ", Desc: " & Err.Description, vbCritical)
    Call DestroyServer
    End

End Sub

Sub LoadItems()

  Dim FileName As String
  Dim i As Long
  Dim f As Long
  Dim Percent As Integer

    Call CheckItems

    For i = 1 To MAX_ITEMS
        Percent = i / MAX_ITEMS * 100
        UpdatePercentage Percent, "Loading items..."

        FileName = App.Path & "\Items\Item" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Item(i)
        Close #f

    Next i

End Sub

Sub LoadMaps()

  Dim FileName As String
  Dim i As Integer
  Dim f As Integer
  Dim Percent As Integer

    If 0 + GetVar(App.Path & "\Data.ini", "CONFIG", "NonToScroll") = 1 Then
        For i = 1 To MAX_MAPS
            Call ClearMapScroll(i)
        Next i
        Call PutVar(App.Path & "\Data.ini", "CONFIG", "NonToScroll", 0)
    End If

    Call CheckMaps

    f = FreeFile
    For i = 1 To MAX_MAPS
        Percent = i / MAX_MAPS * 100
        UpdatePercentage Percent, "Loading maps..."

        Open App.Path & "\maps\map" & i & ".dat" For Binary As #f
            Get #f, , Map(i)
        Close #f

    Next i

End Sub

Sub LoadNpcs()

  Dim i As Integer
  Dim f As Long
  Dim Percent As Integer
  
    Call CheckNpcs
    
    f = FreeFile
    For i = 1 To MAX_NPCS
        Percent = i / MAX_NPCS * 100
        UpdatePercentage Percent, "Loading npcs..."

        Open App.Path & "\npcs\npc" & i & ".dat" For Binary As #f
        Get #f, , Npc(i)
        Close #f

    Next i

End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)

    On Error GoTo PlayerErr
  Dim f As Long 'File
  Dim i As Integer
  Dim FileName As String

    Call ClearPlayer(index)

    FileName = App.Path & "\accounts\" & Trim(Name) & ".ini"

    If FileExist("\accounts\" & Trim(Name) & ".ini") Then
        Call LoadPlayerFromINI(index, Name)
        'Delete the old file
        Kill FileName
     Else
        'Load the account settings
        FileName = App.Path & "\accounts\" & Trim$(Name) & "_info.ini"

        Player(index).Login = Name
        Player(index).Password = GetVar(FileName, "ACCESS", "Password")
        Player(index).Email = GetVar(FileName, "ACCESS", "Email")

        'Load the .dat

        For i = 1 To MAX_CHARS

            FileName = App.Path & "\accounts\" & Trim$(Player(index).Login) & "\char" & i & ".dat"

            f = FreeFile
            Open FileName For Binary As #f
            Get #f, , Player(index).Char(i)
            Close #f

        Next i
    End If

    Exit Sub

PlayerErr:
    Call MsgBox("Error loading player " & index & ". Make sure all variables are correct!", vbCritical)
    Call DestroyServer
    End

End Sub


' Loads a player from an INI file as opposed to a .dat
Sub LoadPlayerFromINI(ByVal index As Integer, ByVal Name As String)

    On Error GoTo PlayerErr
  Dim FileName As String
  Dim i As Long
  Dim j As Long
  Dim n As Long

    FileName = App.Path & "\accounts\" & Name & ".ini"

    Player(index).Login = GetVar(FileName, "GENERAL", "Login")
    Player(index).Password = GetVar(FileName, "GENERAL", "Password")

    For i = 1 To MAX_CHARS

        ' General
        Player(index).Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
        Player(index).Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
        Player(index).Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))
        Player(index).Char(i).Sprite = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
        Player(index).Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
        Player(index).Char(i).Exp = Val(GetVar(FileName, "CHAR" & i, "Exp"))
        Player(index).Char(i).access = Val(GetVar(FileName, "CHAR" & i, "Access"))
        Player(index).Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
        Player(index).Char(i).Guild = GetVar(FileName, "CHAR" & i, "Guild")
        Player(index).Char(i).Guildaccess = Val(GetVar(FileName, "CHAR" & i, "Guildaccess"))
        Player(index).Char(i).head = Val(GetVar(FileName, "CHAR" & i, "Head"))
        Player(index).Char(i).body = Val(GetVar(FileName, "CHAR" & i, "Body"))
        Player(index).Char(i).leg = Val(GetVar(FileName, "CHAR" & i, "Leg"))

        ' Vitals
        Player(index).Char(i).HP = Val(GetVar(FileName, "CHAR" & i, "HP"))
        Player(index).Char(i).MP = Val(GetVar(FileName, "CHAR" & i, "MP"))
        Player(index).Char(i).SP = Val(GetVar(FileName, "CHAR" & i, "SP"))

        ' Stats
        Player(index).Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
        Player(index).Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
        Player(index).Char(i).Speed = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        Player(index).Char(i).Magi = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
        Player(index).Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))

        ' Worn equipment
        Player(index).Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
        Player(index).Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
        Player(index).Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
        Player(index).Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
        Player(index).Char(i).LegsSlot = Val(GetVar(FileName, "CHAR" & i, "LegsSlot"))
        Player(index).Char(i).RingSlot = Val(GetVar(FileName, "CHAR" & i, "RingSlot"))
        Player(index).Char(i).NecklaceSlot = Val(GetVar(FileName, "CHAR" & i, "NecklaceSlot"))

        'sprite stuff
        Player(index).Char(i).head = Val(GetVar(FileName, "CHAR" & i, "Head"))
        Player(index).Char(i).body = Val(GetVar(FileName, "CHAR" & i, "Body"))
        Player(index).Char(i).leg = Val(GetVar(FileName, "CHAR" & i, "Leg"))

        ' Paperdoll

        If GetVar(FileName, "CHAR" & i, "Paperdoll") = "" Then
            Call PutVar(FileName, "CHAR" & i, "Paperdoll", 1)
            Player(index).Char(i).Paperdoll = 1
         Else
            Player(index).Char(i).Paperdoll = Val(GetVar(FileName, "CHAR" & i, "Paperdoll"))
        End If

        'skill

        For j = 1 To MAX_SKILLS
            Player(index).Char(i).SkillLvl(j) = Val(GetVar(FileName, "CHAR" & i, "Skill" & j & "lvl"))
            Player(index).Char(i).SkillExp(j) = Val(GetVar(FileName, "CHAR" & i, "Skill" & j & "exp"))
        Next j

        ' Position
        Player(index).Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
        Player(index).Char(i).X = Val(GetVar(FileName, "CHAR" & i, "X"))
        Player(index).Char(i).Y = Val(GetVar(FileName, "CHAR" & i, "Y"))
        Player(index).Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))

        ' Check to make sure that they aren't on map 0, if so reset'm

        If Player(index).Char(i).Map = 0 Then
            Player(index).Char(i).Map = START_MAP
            Player(index).Char(i).X = START_X
            Player(index).Char(i).Y = START_Y
        End If

        ' Inventory

        For n = 1 To MAX_INV
            Player(index).Char(i).Inv(n).num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
            Player(index).Char(i).Inv(n).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
            Player(index).Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
        Next n

        ' Spells

        For n = 1 To MAX_PLAYER_SPELLS
            Player(index).Char(i).Spell(n) = Val(GetVar(FileName, "CHAR" & i, "Spell" & n))
        Next n

        FileName = App.Path & "\banks\" & Name & ".ini"
        ' Bank

        For n = 1 To MAX_BANK
            Player(index).Char(i).Bank(n).num = Val(GetVar(FileName, "CHAR" & i, "BankItemNum" & n))
            Player(index).Char(i).Bank(n).Value = Val(GetVar(FileName, "CHAR" & i, "BankItemVal" & n))
            Player(index).Char(i).Bank(n).Dur = Val(GetVar(FileName, "CHAR" & i, "BankItemDur" & n))
        Next n

    Next i
    Exit Sub

PlayerErr:
    Call MsgBox("Error loading player " & i & ". Make sure all variables are correct!", vbCritical)
    Call DestroyServer
    End

End Sub

Sub LoadQuests()

  Dim i As Long
  Dim f As Long
  Dim Percent As Integer
  
    Call CheckQuests
    
    f = FreeFile
    For i = 1 To MAX_QUESTS
        Percent = i / MAX_QUESTS * 100
        UpdatePercentage Percent, "Loading quests..."

        Open App.Path & "\Quests\Quest" & i & ".dat" For Binary As #f
        Get #f, , Quest(i)
        Close #f

    Next i

End Sub

Sub LoadShops()

  Dim i As Long
  Dim f As Long
  Dim Percent As Integer
  
    Call CheckShops

    f = FreeFile
    For i = 1 To MAX_SHOPS
        Percent = i / MAX_SHOPS * 100
        UpdatePercentage Percent, "Loading shops..."

        Open App.Path & "\shops\shop" & i & ".dat" For Binary As #f
        Get #f, , Shop(i)
        Close #f

    Next i

End Sub

Sub LoadSkills()

  Dim i As Long
  Dim f As Long
  Dim Percent As Integer

    Call CheckSkills
    
    f = FreeFile
    For i = 1 To MAX_SKILLS
        Percent = i / MAX_SKILLS * 100
        UpdatePercentage Percent, "Loading skills..."

        Open App.Path & "\Skills\Skill" & i & ".dat" For Binary As #f
        Get #f, , skill(i)
        Close #f

    Next i

End Sub

Sub LoadSpells()

  Dim i As Long
  Dim f As Long
  Dim Percent As Integer

    Call CheckSpells
    
    f = FreeFile
    For i = 1 To MAX_SPELLS
        Percent = i / MAX_SPELLS * 100
        UpdatePercentage Percent, "Loading spells..."
        
        Open App.Path & "\spells\spells" & i & ".dat" For Binary As #f
        Get #f, , Spell(i)
        Close #f

    Next i

End Sub

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean

  Dim FileName As String
  Dim RightPassword As String

    PasswordOK = False

    If AccountExist(Name) Then
        'Since we're using the new character save/load we have to check both ways

        If FileExist("\accounts\" & Trim$(Name) & "_info.ini") Then

            FileName = App.Path & "\accounts\" & Trim$(Name) & "_info.ini"
            RightPassword = GetVar(FileName, "ACCESS", "Password")

            If Trim$(Password) = Trim(RightPassword) Then
                PasswordOK = True
            End If

         Else

            FileName = App.Path & "\accounts\" & Trim$(Name) & ".ini"
            RightPassword = GetVar(FileName, "GENERAL", "Password")

            If Trim$(Password) = Trim$(RightPassword) Then
                PasswordOK = True
            End If

        End If
    End If

End Function


'---------------------------------------------------------------------------------------
' Procedure : PutVar
' Purpose   : Writes a file to an INI file
'---------------------------------------------------------------------------------------
Sub PutVar(File As String, Header As String, Var As String, Value As String)

    On Error GoTo PutVar_Error

    Call WritePrivateProfileString(Header, Var, Value, File)

    On Error GoTo 0
    Exit Sub

PutVar_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"

End Sub

Sub SaveAllPlayersOnline()

  Dim i As Integer

    For i = 1 To MAX_PLAYERS

        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If

    Next i

End Sub

Sub SaveArrow(ByVal ArrowNum As Long)

  Dim FileName As String

    FileName = App.Path & "\Arrows.ini"

    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowName", Trim(Arrows(ArrowNum).Name))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowAmount", Val(Arrows(ArrowNum).Amount))

End Sub

Sub SaveClasses()

  Dim FileName As String
  Dim i As Long
  Dim Percent As Integer

    FileName = App.Path & "\Classes\info.ini"

    If Not FileExist("Classes\info.ini") Then
        Call PutVar(FileName, "INFO", "MaxClasses", 3)
        Call PutVar(FileName, "INFO", "MaxSkills", 25)
        Call PutVar(FileName, "INFO", "StatPoints", 0)
        Call PutVar(FileName, "INFO", "SkillPoints", 0)
    End If

    For i = 0 To MAX_CLASSES
        Percent = i / MAX_CLASSES * 100
        UpdatePercentage Percent, "Saving classes..."
        
        FileName = App.Path & "\Classes\Class" & i & ".ini"

        If Not FileExist("Classes\Class" & i & ".ini") Then
            Call PutVar(FileName, "CLASS", "Name", Trim(Class(i).Name))
            Call PutVar(FileName, "CLASS", "MaleSprite", STR(Class(i).MaleSprite))
            Call PutVar(FileName, "CLASS", "FemaleSprite", STR(Class(i).FemaleSprite))
            Call PutVar(FileName, "CLASS", "STR", STR(Class(i).STR))
            Call PutVar(FileName, "CLASS", "DEF", STR(Class(i).DEF))
            Call PutVar(FileName, "CLASS", "SPEED", STR(Class(i).Speed))
            Call PutVar(FileName, "CLASS", "MAGI", STR(Class(i).Magi))
            Call PutVar(FileName, "CLASS", "MAP", STR(Class(i).Map))
            Call PutVar(FileName, "CLASS", "X", STR(Class(i).X))
            Call PutVar(FileName, "CLASS", "Y", STR(Class(i).Y))
            Call PutVar(FileName, "CLASS", "Locked", STR(Class(i).locked))
        End If

    Next i

End Sub

Sub SaveClasses2()

  Dim FileName As String
  Dim i As Long
  Dim Percent As Integer
  
    FileName = App.Path & "\FirstClassAdvancement.ini"

    For i = 0 To MAX_CLASSES
        Percent = i / MAX_CLASSES * 100
        UpdatePercentage Percent, "Saving first class advancement..."
        
        Call PutVar(FileName, "CLASS" & i, "Name", Trim(Class2(i).Name))
        Call PutVar(FileName, "CLASS" & i, "AdvanceFrom", STR(Class2(i).AdvanceFrom))
        Call PutVar(FileName, "CLASS" & i, "LevelReq", STR(Class2(i).LevelReq))
        Call PutVar(FileName, "CLASS" & i, "MaleSprite", STR(Class2(i).MaleSprite))
        Call PutVar(FileName, "CLASS" & i, "FemaleSprite", STR(Class2(i).FemaleSprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class2(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class2(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class2(i).Speed))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class2(i).Magi))
    Next i

End Sub

Sub Saveclasses3()

  Dim FileName As String
  Dim i As Long
  Dim Percent As Integer
  
    FileName = App.Path & "\SecondClassAdvancement.ini"

    For i = 0 To MAX_CLASSES
        Percent = i / MAX_CLASSES * 100
        UpdatePercentage Percent, "Saving second class advancement..."
        
        Call PutVar(FileName, "CLASS" & i, "Name", Trim(Class3(i).Name))
        Call PutVar(FileName, "CLASS" & i, "AdvanceFrom", STR(Class3(i).AdvanceFrom))
        Call PutVar(FileName, "CLASS" & i, "MaleSprite", STR(Class3(i).MaleSprite))
        Call PutVar(FileName, "CLASS" & i, "FemaleSprite", STR(Class3(i).FemaleSprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class3(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class3(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class3(i).Speed))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class3(i).Magi))
    Next i

End Sub

Sub SaveElement(ByVal ElementNum As Long)

  Dim FileName As String

    FileName = App.Path & "\elements.ini"

    Call PutVar(FileName, "ELEMENTS", "ElementName" & ElementNum, Trim(Element(ElementNum).Name))
    Call PutVar(FileName, "ELEMENTS", "ElementStrong" & ElementNum, Val(Element(ElementNum).Strong))
    Call PutVar(FileName, "ELEMENTS", "ElementWeak" & ElementNum, Val(Element(ElementNum).Weak))

End Sub

Sub SaveEmoticon(ByVal EmoNum As Long)

  Dim FileName As String

    FileName = App.Path & "\emoticons.ini"

    Call PutVar(FileName, "EMOTICONS", "EmoticonC" & EmoNum, Trim(Emoticons(EmoNum).Command))
    Call PutVar(FileName, "EMOTICONS", "Emoticon" & EmoNum, Val(Emoticons(EmoNum).Pic))

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
  Dim Percent As Integer
  
    Call SetStatus("Saving items... ")

    For i = 1 To MAX_ITEMS

        If Not FileExist("items\item" & i & ".dat") Then
            Percent = i / MAX_ITEMS * 100
            UpdatePercentage Percent, "Saving items..."
            
            Call SaveItem(i)
            
        End If

    Next i

End Sub

Sub SaveLogs()

  Dim FileName As String
  Dim i As String
  Dim c As String

    On Error Resume Next
    If LCase(Dir(App.Path & "\logs", vbDirectory)) <> "logs" Then
        Call MkDir(App.Path & "\Logs")
    End If

    c = Time
    c = Replace(c, ":", "-", 1, -1, vbTextCompare)
    c = Replace(c, ":", "-", 1, -1, vbTextCompare)
    c = Replace(c, ":", "-", 1, -1, vbTextCompare)

    i = Date

    If LCase(Dir(App.Path & "\logs\" & i, vbDirectory)) <> i Then
        Call MkDir(App.Path & "\Logs\" & i & "\")
    End If

    Call MkDir(App.Path & "\Logs\" & i & "\" & c & "\")

    FileName = App.Path & "\Logs\" & i & "\" & c & "\Main.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(0).text
    Close #1

    FileName = App.Path & "\Logs\" & i & "\" & c & "\Broadcast.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(1).text
    Close #1

    FileName = App.Path & "\Logs\" & i & "\" & c & "\Global.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(2).text
    Close #1

    FileName = App.Path & "\Logs\" & i & "\" & c & "\Map.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(3).text
    Close #1

    FileName = App.Path & "\Logs\" & i & "\" & c & "\Private.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(4).text
    Close #1

    FileName = App.Path & "\Logs\" & i & "\" & c & "\Admin.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(5).text
    Close #1

    FileName = App.Path & "\Logs\" & i & "\" & c & "\Emote.txt"
    Open FileName For Output As #1
    Print #1, frmServer.txtText(6).text
    Close #1

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

Sub SaveNPC(ByVal NpcNum As Long)

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
  Dim Percent As Integer

    Call SetStatus("Saving npcs... ")

    For i = 1 To MAX_NPCS

        If Not FileExist("npcs\npc" & i & ".dat") Then
            Percent = i / MAX_NPCS * 100
            UpdatePercentage Percent, "Saving npcs..."
        
            Call SaveNPC(i)
        End If

    Next i

End Sub

Sub SavePlayer(ByVal index As Long)

  Dim FileName As String
  Dim f As Long 'File
  Dim i As Integer

    'Save login information first
    FileName = App.Path & "\accounts\" & Trim$(Player(index).Login) & "_info.ini"

    Call PutVar(FileName, "ACCESS", "Login", Trim$(Player(index).Login))
    Call PutVar(FileName, "ACCESS", "Password", Trim$(Player(index).Password))
    Call PutVar(FileName, "ACCESS", "Email", Trim$(Player(index).Email))

    'Make the directory

    If LCase(Dir(App.Path & "\accounts\" & Trim$(Player(index).Login), vbDirectory)) <> LCase(Trim$(Player(index).Login)) Then
        Call MkDir(App.Path & "\accounts\" & Trim$(Player(index).Login))
    End If

    'Now save their characters

    For i = 1 To MAX_CHARS
        FileName = App.Path & "\accounts\" & Trim$(Player(index).Login) & "\char" & i & ".dat"

        'Save the character
        f = FreeFile
        Open FileName For Binary As #f
        Put #f, , Player(index).Char(i)
        Close #f

    Next i

End Sub

Sub SaveQuest(ByVal QuestNum As Long)

  Dim FileName As String
  Dim f  As Long

    FileName = App.Path & "\quests\Quest" & QuestNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , Quest(QuestNum)
    Close #f

End Sub

Sub SaveQuests()

  Dim i As Long
  Dim Percent As Integer

    Call SetStatus("Saving quests... ")

    For i = 1 To MAX_QUESTS

        If Not FileExist("Quests\Quest" & i & ".dat") Then
            Percent = i / MAX_QUESTS * 100
            UpdatePercentage Percent, "Saving quest..."
            
            Call SaveQuest(i)
        End If

    Next i

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
  Dim Percent As Integer

    Call SetStatus("Saving shops... ")

    For i = 1 To MAX_SHOPS

        If Not FileExist("shops\shop" & i & ".dat") Then
            Percent = i / MAX_SHOPS * 100
            UpdatePercentage Percent, "Saving shops..."
            
            Call SaveShop(i)
        End If

    Next i

End Sub

Sub SaveSkill(ByVal skillNum As Long)

  Dim FileName As String
  Dim f  As Long

    FileName = App.Path & "\skills\Skill" & skillNum & ".dat"
    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , skill(skillNum)
    Close #f

End Sub

Sub SaveSkills()

  Dim i As Long
  Dim Percent As Integer

    Call SetStatus("Saving skills... ")

    For i = 1 To MAX_SKILLS

        If Not FileExist("Skills\Skill" & i & ".dat") Then
            Percent = i / MAX_SKILLS * 100
            UpdatePercentage Percent, "Saving skills..."
            
            Call SaveSkill(i)
        End If

    Next i

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
  Dim Percent As Integer

    Call SetStatus("Saving spells... ")

    For i = 1 To MAX_SPELLS

        If Not FileExist("spells\spells" & i & ".dat") Then
            Percent = i / MAX_SPELLS * 100
            UpdatePercentage Percent, "Saving spells..."
            
            Call SaveSpell(i)
        End If

    Next i

End Sub

