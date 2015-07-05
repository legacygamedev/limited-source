Attribute VB_Name = "modDatabase"
Option Explicit

' ---------------------------------------------------------------------------------------
' Procedure : GetVar
' Purpose   :  Reads a variable from an INI file
' ---------------------------------------------------------------------------------------
Function GetVar(file As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found

    On Error GoTo GetVar_Error

    szReturn = vbNullString

    sSpaces = Space$(5000)

    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), file)

    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)

    On Error GoTo 0
    Exit Function

GetVar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GetVar of Module modDatabase"
End Function

' ---------------------------------------------------------------------------------------
' Procedure : PutVar
' Purpose   : Writes a file to an INI file
' ---------------------------------------------------------------------------------------
Sub PutVar(file As String, Header As String, Var As String, Value As String)
    On Error GoTo PutVar_Error

    Call WritePrivateProfileString(Header, Var, Value, file)

    On Error GoTo 0
    Exit Sub

PutVar_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PutVar of Module modDatabase"
End Sub

Function FileExists(FileName As String) As Boolean
    On Error GoTo ErrorHandler
    ' get the attributes and ensure that it isn't a directory
    FileExists = (GetAttr(App.Path & "\" & FileName) And vbDirectory) = 0
ErrorHandler:
' if an error occurs, this function returns False
End Function

Function FolderExists(inPath As String) As Boolean
    If LenB(Dir$(inPath, vbDirectory)) = 0 Then
        FolderExists = False
    Else
        FolderExists = True
    End If
End Function

Sub LoadExperience()
    On Error GoTo ExpErr
    Dim FileName As String
    Dim I As Integer

    Call CheckExperience

    FileName = App.Path & "\Experience.ini"

    For I = 1 To MAX_LEVEL
        temp = I / MAX_LEVEL * 100
        Call SetStatus("Loading Experience... " & temp & "%")
        Experience(I) = GetVar(FileName, "EXPERIENCE", "Exp" & I)
    Next I
    Exit Sub

ExpErr:
    Call MsgBox("Error loading EXP for level " & I & ". Make sure Experience.ini has the correct variables! ERR: " & Err.Number & ", Desc: " & Err.Description, vbCritical)
    Call DestroyServer
End Sub

Sub CheckExperience()
    If Not FileExists("Experience.ini") Then
        Dim I As Integer

        For I = 1 To MAX_LEVEL
            temp = I / MAX_LEVEL * 100
            Call SetStatus("Saving Experience... " & temp & "%")
            Call PutVar(App.Path & "\Experience.ini", "EXPERIENCE", "Exp" & I, I * 1500)
        Next I
    End If
End Sub

Sub ClearExperience()
    Dim I As Integer

    For I = 1 To MAX_LEVEL
        Experience(I) = 0
    Next I
End Sub

Sub LoadEmoticon()
    Dim FileName As String
    Dim I As Integer

    Call CheckEmoticon

    FileName = App.Path & "\Emoticons.ini"

    For I = 0 To MAX_EMOTICONS
        temp = I / MAX_EMOTICONS * 100
        Call SetStatus("Loading Emoticons... " & temp & "%")
        Emoticons(I).Pic = GetVar(FileName, "EMOTICONS", "Emoticon" & I)
        Emoticons(I).Command = GetVar(FileName, "EMOTICONS", "EmoticonC" & I)
    Next I
End Sub

Sub SaveEmoticon(ByVal EmoNum As Long)
    Dim FileName As String

    FileName = App.Path & "\Emoticons.ini"

    Call PutVar(FileName, "EMOTICONS", "EmoticonC" & EmoNum, Trim$(Emoticons(EmoNum).Command))
    Call PutVar(FileName, "EMOTICONS", "Emoticon" & EmoNum, Val(Emoticons(EmoNum).Pic))
End Sub

Sub CheckEmoticon()
    If Not FileExists("Emoticons.ini") Then
        Dim I As Integer

        For I = 0 To MAX_EMOTICONS
            temp = I / MAX_LEVEL * 100
            Call SetStatus("Saving emoticons... " & temp & "%")
            Call PutVar(App.Path & "\Emoticons.ini", "EMOTICONS", "Emoticon" & I, 0)
            Call PutVar(App.Path & "\Emoticons.ini", "EMOTICONS", "EmoticonC" & I, vbNullString)
        Next I
    End If
End Sub

Sub ClearEmoticon()
    Dim I As Integer

    For I = 0 To MAX_EMOTICONS
        Emoticons(I).Pic = 0
        Emoticons(I).Command = vbNullString
    Next I
End Sub

Sub LoadElements()
    On Error GoTo ElementErr
    Dim FileName As String
    Dim I As Integer

    Call CheckElements

    FileName = App.Path & "\Elements.ini"

    For I = 0 To MAX_ELEMENTS
        temp = I / MAX_ELEMENTS * 100
        Call SetStatus("Loading elements... " & temp & "%")
        Element(I).Name = GetVar(FileName, "ELEMENTS", "ElementName" & I)
        Element(I).Strong = Val(GetVar(FileName, "ELEMENTS", "ElementStrong" & I))
        Element(I).Weak = Val(GetVar(FileName, "ELEMENTS", "ElementWeak" & I))
    Next I
    Exit Sub

ElementErr:
    Call MsgBox("Error loading element " & I & ". Make sure all the variables in Elements.ini are correct!", vbCritical)
    Call DestroyServer
    End
End Sub

Sub CheckElements()
    If Not FileExists("Elements.ini") Then
        Dim I As Integer

        For I = 0 To MAX_ELEMENTS
            temp = I / MAX_ELEMENTS * 100
            Call SetStatus("Saving elements... " & temp & "%")
            Call PutVar(App.Path & "\Elements.ini", "ELEMENTS", "ElementName" & I, vbNullString)
            Call PutVar(App.Path & "\Elements.ini", "ELEMENTS", "ElementStrong" & I, 0)
            Call PutVar(App.Path & "\Elements.ini", "ELEMENTS", "ElementWeak" & I, 0)
        Next I
    End If
End Sub

Sub SaveElement(ByVal ElementNum As Long)
    Dim FileName As String

    FileName = App.Path & "\Elements.ini"

    Call PutVar(FileName, "ELEMENTS", "ElementName" & ElementNum, Trim$(Element(ElementNum).Name))
    Call PutVar(FileName, "ELEMENTS", "ElementStrong" & ElementNum, Val(Element(ElementNum).Strong))
    Call PutVar(FileName, "ELEMENTS", "ElementWeak" & ElementNum, Val(Element(ElementNum).Weak))
End Sub

Sub SavePlayer(ByVal Index As Long)
    Dim FileName As String
    Dim f As Long 'File
    Dim I As Integer

    On Error Resume Next

    ' Save login information first
    FileName = App.Path & "\Accounts\" & Trim$(Player(Index).Login) & "_Info.ini"

    Call PutVar(FileName, "ACCESS", "Login", Trim$(Player(Index).Login))
    Call PutVar(FileName, "ACCESS", "Password", Trim$(Player(Index).Password))
    Call PutVar(FileName, "ACCESS", "Email", Trim$(Player(Index).Email))

    ' Make the directory
    If LCase$(Dir$(App.Path & "\Accounts\" & Trim$(Player(Index).Login), vbDirectory)) <> LCase$(Trim$(Player(Index).Login)) Then
        Call MkDir(App.Path & "\Accounts\" & Trim$(Player(Index).Login))
    End If

    ' Now save their characters
    For I = 1 To MAX_CHARS
        FileName = App.Path & "\Accounts\" & Trim$(Player(Index).Login) & "\Char" & I & ".dat"

        ' Save the character
        f = FreeFile
        Open FileName For Binary As #f
        Put #f, , Player(Index).Char(I)
        Close #f

    Next I
End Sub

Function ConvertV000(FileName As String) As PlayerRec
    Dim OldRec As V000PlayerRec
    Dim NewRec As PlayerRec
    Dim f As Long
    Dim n As Integer
    
    f = FreeFile
    Open FileName For Binary As #f
        Get #f, , OldRec
    Close #f

        ' General
    NewRec.Name = OldRec.Name
    NewRec.Guild = OldRec.Guild
    NewRec.GuildAccess = OldRec.GuildAccess
    NewRec.Sex = OldRec.Sex
    NewRec.Class = OldRec.Class
    NewRec.Sprite = OldRec.Sprite
    NewRec.LEVEL = OldRec.LEVEL
    NewRec.Exp = OldRec.Exp
    NewRec.Access = OldRec.Access
    NewRec.PK = OldRec.PK

    ' Vitals
    NewRec.HP = OldRec.HP
    NewRec.MP = OldRec.MP
    NewRec.SP = OldRec.SP

    ' Stats
    NewRec.STR = OldRec.STR
    NewRec.DEF = OldRec.DEF
    NewRec.Speed = OldRec.Speed
    NewRec.Magi = OldRec.Magi
    NewRec.POINTS = OldRec.POINTS

    ' Worn equipment
    NewRec.ArmorSlot = OldRec.ArmorSlot
    NewRec.WeaponSlot = OldRec.WeaponSlot
    NewRec.HelmetSlot = OldRec.HelmetSlot
    NewRec.ShieldSlot = OldRec.ShieldSlot
    NewRec.LegsSlot = OldRec.LegsSlot
    NewRec.RingSlot = OldRec.RingSlot
    NewRec.NecklaceSlot = OldRec.NecklaceSlot

    ' Inventory
    For n = 1 To MAX_INV
        NewRec.Inv(n) = OldRec.Inv(n)
    Next n
    For n = 1 To MAX_PLAYER_SPELLS
        NewRec.Spell(n) = OldRec.Spell(n)
    Next n
    For n = 1 To MAX_BANK
        NewRec.Bank(n) = OldRec.Bank(n)
    Next n

    ' Position
    NewRec.Map = OldRec.Map
    NewRec.X = OldRec.X
    NewRec.Y = OldRec.Y
    NewRec.Dir = OldRec.Dir

    NewRec.TargetNPC = OldRec.TargetNPC

    NewRec.Head = OldRec.Head
    NewRec.Body = OldRec.Body
    NewRec.Leg = OldRec.Leg

    NewRec.PAPERDOLL = OldRec.PAPERDOLL

    NewRec.MAXHP = OldRec.MAXHP
    NewRec.MAXMP = OldRec.MAXMP
    NewRec.MAXSP = OldRec.MAXSP


    ' *** add new fields ***

    ' version info
    
    NewRec.Vflag = 128
    NewRec.Ver = 2
    NewRec.SubVer = 8
    NewRec.Rel = 0

    ConvertV000 = NewRec


End Function

Public Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
    Dim f As Long
    Dim I As Integer
    Dim FileName As String

    On Error GoTo PlayerErr

    Call ClearPlayer(Index)

    ' Load the account settings
    FileName = App.Path & "\Accounts\" & Trim$(Name) & "_Info.ini"

    Player(Index).Login = Name
    Player(Index).Password = GetVar(FileName, "ACCESS", "Password")
    Player(Index).Email = GetVar(FileName, "ACCESS", "Email")

    ' Load the .dat
    For I = 1 To MAX_CHARS
        FileName = App.Path & "\Accounts\" & Trim$(Player(Index).Login) & "\Char" & I & ".dat"

        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Player(Index).Char(I)
        Close #f
        
        If Player(Index).Char(I).Vflag <> 128 Then
            Player(Index).Char(I) = ConvertV000(FileName)
        End If
        
    Next I

    Exit Sub

PlayerErr:
    Call MsgBox("Couldn't load index " & Index & " for " & Name & "!", vbCritical)
    Call DestroyServer
End Sub

Function AccountExists(ByVal Name As String) As Boolean
    If FileExists("\Accounts\" & Trim$(Name) & "_Info.ini") Then
        AccountExists = True
    Else
        AccountExists = False
    End If
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim$(Player(Index).Char(CharNum).Name) <> vbNullString Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
    Dim RightPassword As String

    PasswordOK = False

    If AccountExists(Name) Then
        RightPassword = GetVar(App.Path & "\Accounts\" & Trim$(Name) & "_Info.ini", "ACCESS", "Password")

        If Trim$(Password) = Trim$(RightPassword) Then
            PasswordOK = True
        End If
    End If
End Function

Public Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String, ByVal Email As String)
    Dim I As Long

    Player(Index).Login = Name
    Player(Index).Password = Password
    Player(Index).Email = Email

    For I = 1 To MAX_CHARS
        Call ClearChar(Index, I)
    Next I

    Call SavePlayer(Index)

    If ACC_VERIFY = 1 Then
        Call PutVar(App.Path & "\Accounts\" & Trim$(Player(Index).Login) & "_Info.ini", "ACCESS", "verified", 0)
    End If
    
    Call ClearPlayer(Index)
End Sub

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long, ByVal headc As Long, ByVal bodyc As Long, ByVal logc As Long)
    Dim f As Long

    If Trim$(Player(Index).Char(CharNum).Name) = vbNullString Then
        Player(Index).CharNum = CharNum

        Player(Index).Char(CharNum).Name = Name
        Player(Index).Char(CharNum).Sex = Sex
        Player(Index).Char(CharNum).Class = ClassNum

        If Player(Index).Char(CharNum).Sex = SEX_MALE Then
            Player(Index).Char(CharNum).Sprite = ClassData(ClassNum).MaleSprite
        Else
            Player(Index).Char(CharNum).Sprite = ClassData(ClassNum).FemaleSprite
        End If

        Player(Index).Char(CharNum).LEVEL = 1

        Player(Index).Char(CharNum).STR = ClassData(ClassNum).STR
        Player(Index).Char(CharNum).DEF = ClassData(ClassNum).DEF
        Player(Index).Char(CharNum).Speed = ClassData(ClassNum).Speed
        Player(Index).Char(CharNum).Magi = ClassData(ClassNum).Magi

        If ClassData(ClassNum).Map <= 0 Then
            ClassData(ClassNum).Map = 1
        End If
        If ClassData(ClassNum).X < 0 Or ClassData(ClassNum).X > MAX_MAPX Then
            ClassData(ClassNum).X = Int(ClassData(ClassNum).X / 2)
        End If
        If ClassData(ClassNum).Y < 0 Or ClassData(ClassNum).Y > MAX_MAPY Then
            ClassData(ClassNum).Y = Int(ClassData(ClassNum).Y / 2)
        End If
        Player(Index).Char(CharNum).Map = ClassData(ClassNum).Map
        Player(Index).Char(CharNum).X = ClassData(ClassNum).X
        Player(Index).Char(CharNum).Y = ClassData(ClassNum).Y

        Player(Index).Char(CharNum).HP = GetPlayerMaxHP(Index)
        Player(Index).Char(CharNum).MP = GetPlayerMaxMP(Index)
        Player(Index).Char(CharNum).SP = GetPlayerMaxSP(Index)

        Player(Index).Char(CharNum).MAXHP = GetPlayerMaxHP(Index)
        Player(Index).Char(CharNum).MAXMP = GetPlayerMaxMP(Index)
        Player(Index).Char(CharNum).MAXSP = GetPlayerMaxSP(Index)

        Player(Index).Char(CharNum).Head = headc
        Player(Index).Char(CharNum).Body = bodyc
        Player(Index).Char(CharNum).Leg = logc
        
        ' version info
        Player(Index).Char(CharNum).Vflag = 128
        Player(Index).Char(CharNum).Ver = 2
        Player(Index).Char(CharNum).SubVer = 8
        Player(Index).Char(CharNum).Rel = 0

        Player(Index).Char(CharNum).PAPERDOLL = 1

        ' Append name to file
        f = FreeFile
        Open App.Path & "\Accounts\CharList.txt" For Append As #f
        Print #f, Name
        Close #f

        Call SavePlayer(Index)

        Exit Sub
    End If
End Sub

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim f As Long
    Dim S As String

    FindChar = False

    f = FreeFile
    Open App.Path & "\Accounts\CharList.txt" For Input As #f
    Do While Not EOF(f)
        Input #f, S

        If Trim$(LCase$(S)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #f
            Exit Function
        End If
    Loop
    Close #f
End Function

Sub SaveAllPlayersOnline()
    Dim I As Integer

    For I = 1 To MAX_PLAYERS
        If IsPlaying(I) Then
            Call SavePlayer(I)
        End If
    Next I
End Sub

Sub LoadClasses()
    Dim FileName As String
    Dim I As Long

    ' -1 because it is 0 indexed and it needs to include 0
    MAX_CLASSES = -1

    Do While FileExists("\Classes\Class" & I & ".ini")
        MAX_CLASSES = MAX_CLASSES + 1
        I = I + 1
    Loop
    
    If MAX_CLASSES = -1 Then
        MAX_CLASSES = 0
    End If

    ReDim ClassData(0 To MAX_CLASSES) As ClassRec

    Call ClearClasses

    For I = 0 To MAX_CLASSES

        On Error Resume Next ' used if next line tries to divide by 0

        temp = I / MAX_CLASSES * 100

        On Error GoTo ClassErr

        Call SetStatus("Loading classes... " & temp & "%")
        FileName = App.Path & "\Classes\Class" & I & ".ini"

        ' Check if class exists
        If Not FileExists("\Classes\Class" & I & ".ini") Then
            Call PutVar(FileName, "CLASS", "Name", Trim$(ClassData(I).Name))
            Call PutVar(FileName, "CLASS", "MaleSprite", CStr(ClassData(I).MaleSprite))
            Call PutVar(FileName, "CLASS", "FemaleSprite", CStr(ClassData(I).FemaleSprite))
            Call PutVar(FileName, "CLASS", "Description", CStr(ClassData(I).Desc))
            Call PutVar(FileName, "CLASS", "STR", CStr(ClassData(I).STR))
            Call PutVar(FileName, "CLASS", "DEF", CStr(ClassData(I).DEF))
            Call PutVar(FileName, "CLASS", "SPEED", CStr(ClassData(I).Speed))
            Call PutVar(FileName, "CLASS", "MAGI", CStr(ClassData(I).Magi))
            Call PutVar(FileName, "CLASS", "MAP", CStr(ClassData(I).Map))
            Call PutVar(FileName, "CLASS", "X", CStr(ClassData(I).X))
            Call PutVar(FileName, "CLASS", "Y", CStr(ClassData(I).Y))
            Call PutVar(FileName, "CLASS", "Locked", CStr(ClassData(I).Locked))
        End If

        ClassData(I).Name = GetVar(FileName, "CLASS", "Name")
        ClassData(I).MaleSprite = GetVar(FileName, "CLASS", "MaleSprite")
        ClassData(I).FemaleSprite = GetVar(FileName, "CLASS", "FemaleSprite")
        ClassData(I).Desc = GetVar(FileName, "CLASS", "Desc")
        ClassData(I).STR = Val(GetVar(FileName, "CLASS", "STR"))
        ClassData(I).DEF = Val(GetVar(FileName, "CLASS", "DEF"))
        ClassData(I).Speed = Val(GetVar(FileName, "CLASS", "SPEED"))
        ClassData(I).Magi = Val(GetVar(FileName, "CLASS", "MAGI"))
        ClassData(I).Map = Val(GetVar(FileName, "CLASS", "MAP"))
        ClassData(I).X = Val(GetVar(FileName, "CLASS", "X"))
        ClassData(I).Y = Val(GetVar(FileName, "CLASS", "Y"))
        ClassData(I).Locked = Val(GetVar(FileName, "CLASS", "Locked"))
    Next I
    Exit Sub

ClassErr:
    Call MsgBox("Error loading class " & I & ". Check that all the variables in your class files exist!")
    Call DestroyServer
    End
End Sub

Sub SaveClasses()
    Dim FileName As String
    Dim I As Long

    For I = 0 To MAX_CLASSES
        On Error Resume Next ' if MAX_CLASSES is 0
    
        temp = I / MAX_CLASSES * 100
        Call SetStatus("Saving classes... " & temp & "%")
        FileName = App.Path & "\Classes\Class" & I & ".ini"
        If Not FileExists("Classes\Class" & I & ".ini") Then
            Call PutVar(FileName, "CLASS", "Name", Trim$(ClassData(I).Name))
            Call PutVar(FileName, "CLASS", "MaleSprite", CStr(ClassData(I).MaleSprite))
            Call PutVar(FileName, "CLASS", "FemaleSprite", CStr(ClassData(I).FemaleSprite))
            Call PutVar(FileName, "CLASS", "STR", CStr(ClassData(I).STR))
            Call PutVar(FileName, "CLASS", "DEF", CStr(ClassData(I).DEF))
            Call PutVar(FileName, "CLASS", "SPEED", CStr(ClassData(I).Speed))
            Call PutVar(FileName, "CLASS", "MAGI", CStr(ClassData(I).Magi))
            Call PutVar(FileName, "CLASS", "MAP", CStr(ClassData(I).Map))
            Call PutVar(FileName, "CLASS", "X", CStr(ClassData(I).X))
            Call PutVar(FileName, "CLASS", "Y", CStr(ClassData(I).Y))
            Call PutVar(FileName, "CLASS", "Locked", CStr(ClassData(I).Locked))
        End If
    Next I
End Sub

Sub SaveItems()
    Dim I As Long

    Call SetStatus("Saving items... ")
    For I = 1 To MAX_ITEMS
        If Not FileExists("items\item" & I & ".dat") Then
            temp = I / MAX_ITEMS * 100
            Call SetStatus("Saving items... " & temp & "%")
            Call SaveItem(I)
        End If
    Next I
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

Sub LoadItems()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckItems

    For I = 1 To MAX_ITEMS
        temp = I / MAX_ITEMS * 100
        Call SetStatus("Loading items... " & temp & "%")

        FileName = App.Path & "\Items\Item" & I & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Item(I)
        Close #f

    Next I
End Sub

Sub CheckItems()
    Call SaveItems
End Sub

Sub SaveShops()
    Dim I As Long

    Call SetStatus("Saving shops... ")
    For I = 1 To MAX_SHOPS
        If Not FileExists("shops\shop" & I & ".dat") Then
            temp = I / MAX_SHOPS * 100
            Call SetStatus("Saving shops... " & temp & "%")
            Call SaveShop(I)
        End If
    Next I
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

Sub LoadShops()
    Dim FileName As String
    Dim I As Long, f As Long

    Call CheckShops

    For I = 1 To MAX_SHOPS
        temp = I / MAX_SHOPS * 100
        Call SetStatus("Loading shops... " & temp & "%")
        FileName = App.Path & "\shops\shop" & I & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Shop(I)
        Close #f

    Next I
End Sub

Sub CheckShops()
    Call SaveShops
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
    Dim I As Long

    Call SetStatus("Saving spells... ")
    For I = 1 To MAX_SPELLS
        If Not FileExists("spells\spells" & I & ".dat") Then
            temp = I / MAX_SPELLS * 100
            Call SetStatus("Saving spells... " & temp & "%")
            Call SaveSpell(I)
        End If
    Next I
End Sub

Sub LoadSpells()
    Dim FileName As String
    Dim I As Long
    Dim f As Long

    Call CheckSpells

    For I = 1 To MAX_SPELLS
        temp = I / MAX_SPELLS * 100
        Call SetStatus("Loading spells... " & temp & "%")

        FileName = App.Path & "\spells\spells" & I & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Spell(I)
        Close #f

    Next I
End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub SaveNpcs()
    Dim I As Long

    Call SetStatus("Saving npcs... ")

    For I = 1 To MAX_NPCS
        If Not FileExists("npcs\npc" & I & ".dat") Then
            temp = I / MAX_NPCS * 100
            Call SetStatus("Saving npcs... " & temp & "%")
            Call SaveNpc(I)
        End If
    Next I
End Sub

Sub SaveNpc(ByVal NPCnum As Long)
    Dim FileName As String
    Dim f As Long
    FileName = App.Path & "\npcs\npc" & NPCnum & ".dat"

    f = FreeFile
    Open FileName For Binary As #f
    Put #f, , NPC(NPCnum)
    Close #f
End Sub

Sub LoadNpcs()
    Dim FileName As String
    Dim I As Integer
    Dim f As Long

    Call CheckNpcs

    For I = 1 To MAX_NPCS
        temp = I / MAX_NPCS * 100
        Call SetStatus("Loading npcs... " & temp & "%")
        FileName = App.Path & "\npcs\npc" & I & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , NPC(I)
        Close #f

    Next I
End Sub

Sub CheckNpcs()
    Call SaveNpcs
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

Sub LoadMaps()
    Dim FileName As String
    Dim I As Integer
    Dim f As Integer

    Call CheckMaps

    For I = 1 To MAX_MAPS
        temp = I / MAX_MAPS * 100
        Call SetStatus("Loading maps... " & temp & "%")
        FileName = App.Path & "\maps\map" & I & ".dat"

        f = FreeFile
        Open FileName For Binary As #f
        Get #f, , Map(I)
        Close #f

    Next I

End Sub

Sub CheckMaps()
    Dim FileName As String
    Dim I As Integer

    Call ClearMaps

    For I = 1 To MAX_MAPS
        FileName = "maps\map" & I & ".dat"

        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExists(FileName) Then
            temp = I / MAX_MAPS * 100
            Call SetStatus("Saving maps... " & temp & "%")
            Call SaveMap(I)
        End If
    Next I
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
    Dim FileName As String
    Dim FileID As Long

    FileName = App.Path & "\BanList.txt"

    FileID = FreeFile
    Open FileName For Append As #FileID
        Print #FileID, GetPlayerIP(BanPlayerIndex)
    Close #FileID

    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", WHITE)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Public Sub AddLog(ByVal text As String, ByVal FN As String)
    Dim FileName As String
    Dim FileID As Long

    If ServerLog Then
        FileName = App.Path & "\" & FN

        If FileExists(FN) Then
            FileID = FreeFile
            Open FileName For Output As #FileID
            Print #FileID, Time & ": " & text
            Close #FileID
        End If
    End If
End Sub

Public Sub PacketDump(ByVal LogMessage As String, ByVal FileName As String)

   Dim FileID As Integer
  
   ' Check if we're logging.
   If Not ServerLog Then Exit Sub
  
   ' Build the full file name.
   FileName = App.Path & "\" & FileName
  
   ' Get an available file number to write to.
   FileID = FreeFile()
  
   ' Open up the requested file and write the data.
   Open FileName For Append As #FileID
       Print #FileID, Time & ": " & LogMessage
   Close #FileID
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long, f2 As Long
    Dim S As String

    Call FileCopy(App.Path & "\Accounts\CharList.txt", App.Path & "\Accounts\chartemp.txt")

    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\Accounts\chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\Accounts\CharList.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, S
        If Trim$(LCase$(S)) <> Trim$(LCase$(Name)) Then
            Print #f2, S
        End If
    Loop

    Close #f1
    Close #f2

    Call Kill(App.Path & "\Accounts\chartemp.txt")
End Sub

Sub BanByServer(ByVal Index As Long, ByVal Reason As String)
    Dim FileName As String
    Dim FileID As Long
    
    If IsPlaying(Index) Then
        FileName = App.Path & "\BanList.txt"

        FileID = FreeFile
        Open FileName For Append As #FileID
            Print #FileID, GetPlayerIP(Index)
        Close #FileID

        If LenB(Reason) <> 0 Then
            Call GlobalMsg(GetPlayerName(Index) & " has been banned by the server! Reason(" & Reason & ")", WHITE)
            Call AddLog("The server has banned " & GetPlayerName(Index) & ". Reason(" & Reason & ")", ADMIN_LOG)
            Call AlertMsg(Index, "You have been banned by the server!  Reason(" & Reason & ")")
        Else
            Call GlobalMsg(GetPlayerName(Index) & " has been banned by the server!", WHITE)
            Call AddLog("The server has banned " & GetPlayerName(Index) & ".", ADMIN_LOG)
            Call AlertMsg(Index, "You have been banned by the server!")
        End If
    End If
End Sub

Sub SaveLogs()
    Dim FileName As String
    Dim CurDate As String
    Dim CurTime As String
    Dim FileID As Integer

    On Error Resume Next

    If Not FolderExists(App.Path & "\Logs") Then
        Call MkDir(App.Path & "\Logs")
    End If

    CurDate = Replace(Date, "/", "-")

    CurTime = Replace(Time, ":", "-")

    If Not FolderExists(App.Path & "\Logs\" & CurDate) Then
        Call MkDir(App.Path & "\Logs\" & CurDate)
    End If

    Call MkDir(App.Path & "\Logs\" & CurDate & "\" & CurTime)

    FileID = FreeFile

    FileName = App.Path & "\Logs\" & CurDate & "\" & CurTime & "\main.ess"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(0).text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "\" & CurTime & "\Broadcast.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(1).text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "\" & CurTime & "\Global.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(2).text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "\" & CurTime & "\Map.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(3).text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "\" & CurTime & "\Private.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(4).text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "\" & CurTime & "\Admin.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(5).text
    Close #FileID

    FileName = App.Path & "\Logs\" & CurDate & "\" & CurTime & "\Emote.txt"
    Open FileName For Output As #FileID
        Print #FileID, frmServer.txtText(6).text
    Close #FileID
    
    Call TextAdd(frmServer.txtText(0), "The chat logs were successfully saved!", True)
End Sub

Sub LoadArrows()
    Dim FileName As String
    Dim I As Long

    Call CheckArrows

    FileName = App.Path & "\Arrows.ini"

    For I = 1 To MAX_ARROWS
        temp = I / MAX_ARROWS * 100
        Call SetStatus("Loading Arrows... " & temp & "%")
        Arrows(I).Name = GetVar(FileName, "Arrow" & I, "ArrowName")
        Arrows(I).Pic = GetVar(FileName, "Arrow" & I, "ArrowPic")
        Arrows(I).Range = GetVar(FileName, "Arrow" & I, "ArrowRange")
        Arrows(I).Amount = GetVar(FileName, "Arrow" & I, "ArrowAmount")

    Next I
End Sub

Sub CheckArrows()
    If Not FileExists("Arrows.ini") Then
        Dim I As Long

        For I = 1 To MAX_ARROWS
            temp = I / MAX_ARROWS * 100
            Call SetStatus("Saving arrows... " & temp & "%")

            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowName", vbNullString)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowPic", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowRange", 0)
            Call PutVar(App.Path & "\Arrows.ini", "Arrow" & I, "ArrowAmount", 0)
        Next I
    End If
End Sub

Sub ClearArrows()
    Dim I As Long

    For I = 1 To MAX_ARROWS
        Arrows(I).Name = vbNullString
        Arrows(I).Pic = 0
        Arrows(I).Range = 0
        Arrows(I).Amount = 0
    Next I
End Sub

Sub SaveArrow(ByVal ArrowNum As Long)
    Dim FileName As String

    FileName = App.Path & "\Arrows.ini"

    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowName", Trim$(Arrows(ArrowNum).Name))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowPic", Val(Arrows(ArrowNum).Pic))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowRange", Val(Arrows(ArrowNum).Range))
    Call PutVar(FileName, "Arrow" & ArrowNum, "ArrowAmount", Val(Arrows(ArrowNum).Amount))
End Sub

Sub ClearTempTile()
    Dim I As Long, Y As Long, X As Long

    For I = 1 To MAX_MAPS
        TempTile(I).DoorTimer = 0

        For Y = 0 To MAX_MAPY
            For X = 0 To MAX_MAPX
                TempTile(I).DoorOpen(X, Y) = NO
            Next X
        Next Y
    Next I
End Sub

Sub ClearClasses()
    Dim I As Long

    For I = 0 To MAX_CLASSES
        ClassData(I).Name = vbNullString
        ClassData(I).AdvanceFrom = 0
        ClassData(I).LevelReq = 0
        ClassData(I).Type = 1
        ClassData(I).STR = 0
        ClassData(I).DEF = 0
        ClassData(I).Speed = 0
        ClassData(I).Magi = 0
        ClassData(I).FemaleSprite = 0
        ClassData(I).MaleSprite = 0
        ClassData(I).Desc = vbNullString
        ClassData(I).Map = 0
        ClassData(I).X = 0
        ClassData(I).Y = 0
    Next I
End Sub

Sub ClearPlayer(ByVal Index As Long)
    Dim I As Long
    Dim n As Long

    Player(Index).Login = vbNullString
    Player(Index).Password = vbNullString
    For I = 1 To MAX_CHARS
        Player(Index).Char(I).Name = vbNullString
        Player(Index).Char(I).Class = 0
        Player(Index).Char(I).LEVEL = 0
        Player(Index).Char(I).Sprite = 0
        Player(Index).Char(I).Exp = 0
        Player(Index).Char(I).Access = 0
        Player(Index).Char(I).PK = NO
        Player(Index).Char(I).POINTS = 0
        Player(Index).Char(I).Guild = vbNullString

        Player(Index).Char(I).HP = 0
        Player(Index).Char(I).MP = 0
        Player(Index).Char(I).SP = 0

        Player(Index).Char(I).MAXHP = 0
        Player(Index).Char(I).MAXMP = 0
        Player(Index).Char(I).MAXSP = 0

        Player(Index).Char(I).STR = 0
        Player(Index).Char(I).DEF = 0
        Player(Index).Char(I).Speed = 0
        Player(Index).Char(I).Magi = 0

        For n = 1 To MAX_INV
            Player(Index).Char(I).Inv(n).num = 0
            Player(Index).Char(I).Inv(n).Value = 0
            Player(Index).Char(I).Inv(n).Dur = 0
        Next n
        For n = 1 To MAX_BANK
            Player(Index).Char(I).Bank(n).num = 0
            Player(Index).Char(I).Bank(n).Value = 0
            Player(Index).Char(I).Bank(n).Dur = 0
        Next n
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(I).Spell(n) = 0
        Next n

        Player(Index).Char(I).ArmorSlot = 0
        Player(Index).Char(I).WeaponSlot = 0
        Player(Index).Char(I).HelmetSlot = 0
        Player(Index).Char(I).ShieldSlot = 0
        Player(Index).Char(I).LegsSlot = 0
        Player(Index).Char(I).RingSlot = 0
        Player(Index).Char(I).NecklaceSlot = 0

        Player(Index).Char(I).Map = 0
        Player(Index).Char(I).X = 0
        Player(Index).Char(I).Y = 0
        Player(Index).Char(I).Dir = 0

        Player(Index).Locked = False
        Player(Index).LockedSpells = False
        Player(Index).LockedItems = False
        Player(Index).LockedAttack = False

        ' Temporary vars
        Player(Index).Buffer = vbNullString
        Player(Index).IncBuffer = vbNullString
        Player(Index).CharNum = 0
        Player(Index).InGame = False
        Player(Index).AttackTimer = 0
        Player(Index).DataTimer = 0
        Player(Index).DataBytes = 0
        Player(Index).DataPackets = 0
        Player(Index).PartyPlayer = 0
        Player(Index).InParty = False
        Player(Index).Target = 0
        Player(Index).TargetType = 0
        Player(Index).CastedSpell = NO
        Player(Index).PartyStarter = NO
        Player(Index).GettingMap = NO
        Player(Index).Emoticon = -1
        Player(Index).InTrade = False
        Player(Index).TradePlayer = 0
        Player(Index).TradeOk = 0
        Player(Index).TradeItemMax = 0
        Player(Index).TradeItemMax2 = 0
        For n = 1 To MAX_PLAYER_TRADES
            Player(Index).Trading(n).InvName = vbNullString
            Player(Index).Trading(n).InvNum = 0
        Next n
        Player(Index).ChatPlayer = 0
    Next I

End Sub

Sub ClearChar(ByVal Index As Long, ByVal CharNum As Long)
    Dim n As Long
    
    ' version info
    Player(Index).Char(CharNum).Vflag = 128
    Player(Index).Char(CharNum).Ver = 2
    Player(Index).Char(CharNum).SubVer = 8
    Player(Index).Char(CharNum).Rel = 0

    Player(Index).Char(CharNum).Name = vbNullString
    Player(Index).Char(CharNum).Class = 0
    Player(Index).Char(CharNum).Sprite = 0
    Player(Index).Char(CharNum).LEVEL = 0
    Player(Index).Char(CharNum).Exp = 0
    Player(Index).Char(CharNum).Access = 0
    Player(Index).Char(CharNum).PK = NO
    Player(Index).Char(CharNum).POINTS = 0
    Player(Index).Char(CharNum).Guild = vbNullString

    Player(Index).Char(CharNum).HP = 0
    Player(Index).Char(CharNum).MP = 0
    Player(Index).Char(CharNum).SP = 0

    Player(Index).Char(CharNum).MAXHP = 0
    Player(Index).Char(CharNum).MAXMP = 0
    Player(Index).Char(CharNum).MAXSP = 0

    Player(Index).Char(CharNum).STR = 0
    Player(Index).Char(CharNum).DEF = 0
    Player(Index).Char(CharNum).Speed = 0
    Player(Index).Char(CharNum).Magi = 0

    For n = 1 To MAX_INV
        Player(Index).Char(CharNum).Inv(n).num = 0
        Player(Index).Char(CharNum).Inv(n).Value = 0
        Player(Index).Char(CharNum).Inv(n).Dur = 0
    Next n
    For n = 1 To MAX_BANK
        Player(Index).Char(CharNum).Bank(n).num = 0
        Player(Index).Char(CharNum).Bank(n).Value = 0
        Player(Index).Char(CharNum).Bank(n).Dur = 0
    Next n
    For n = 1 To MAX_PLAYER_SPELLS
        Player(Index).Char(CharNum).Spell(n) = 0
    Next n

    Player(Index).Char(CharNum).ArmorSlot = 0
    Player(Index).Char(CharNum).WeaponSlot = 0
    Player(Index).Char(CharNum).HelmetSlot = 0
    Player(Index).Char(CharNum).ShieldSlot = 0
    Player(Index).Char(CharNum).LegsSlot = 0
    Player(Index).Char(CharNum).RingSlot = 0
    Player(Index).Char(CharNum).NecklaceSlot = 0

    Player(Index).Char(CharNum).Map = 0
    Player(Index).Char(CharNum).X = 0
    Player(Index).Char(CharNum).Y = 0
    Player(Index).Char(CharNum).Dir = 0
End Sub

Sub ClearItem(ByVal Index As Long)
    Item(Index).Name = vbNullString
    Item(Index).Desc = vbNullString

    Item(Index).Type = 0
    Item(Index).Data1 = 0
    Item(Index).Data2 = 0
    Item(Index).Data3 = 0
    Item(Index).StrReq = 0
    Item(Index).DefReq = 0
    Item(Index).SpeedReq = 0
    Item(Index).MagicReq = 0
    Item(Index).ClassReq = -1
    Item(Index).AccessReq = 0

    Item(Index).addHP = 0
    Item(Index).addMP = 0
    Item(Index).addSP = 0
    Item(Index).AddStr = 0
    Item(Index).AddDef = 0
    Item(Index).AddMagi = 0
    Item(Index).AddSpeed = 0
    Item(Index).AddEXP = 0
    Item(Index).AttackSpeed = 1000
    Item(Index).Price = 0
    Item(Index).Stackable = 0
    Item(Index).Bound = 0
End Sub

Sub ClearItems()
    Dim I As Long

    For I = 1 To MAX_ITEMS
        Call ClearItem(I)
    Next I
End Sub

Sub ClearNpc(ByVal Index As Long)
    Dim I As Long
    NPC(Index).Name = vbNullString
    NPC(Index).AttackSay = vbNullString
    NPC(Index).Sprite = 0
    NPC(Index).SpawnSecs = 0
    NPC(Index).Behavior = 0
    NPC(Index).Range = 0
    NPC(Index).STR = 0
    NPC(Index).DEF = 0
    NPC(Index).Speed = 0
    NPC(Index).Magi = 0
    NPC(Index).Big = 0
    NPC(Index).MAXHP = 0
    NPC(Index).Exp = 0
    NPC(Index).SpawnTime = 0
    NPC(Index).Element = 0

    For I = 1 To MAX_NPC_DROPS
        NPC(Index).ItemNPC(I).Chance = 0
        NPC(Index).ItemNPC(I).ItemNum = 0
        NPC(Index).ItemNPC(I).ItemValue = 0
    Next I

End Sub

Sub ClearNpcs()
    Dim I As Long

    For I = 1 To MAX_NPCS
        Call ClearNpc(I)
    Next I
End Sub

Sub ClearMapItem(ByVal Index As Long, ByVal MapNum As Long)
    MapItem(MapNum, Index).num = 0
    MapItem(MapNum, Index).Value = 0
    MapItem(MapNum, Index).Dur = 0
    MapItem(MapNum, Index).X = 0
    MapItem(MapNum, Index).Y = 0
End Sub

Sub ClearMapItems()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(X, Y)
        Next X
    Next Y
End Sub

Sub ClearMapNpc(ByVal Index As Long, ByVal MapNum As Long)
    MapNPC(MapNum, Index).num = 0
    MapNPC(MapNum, Index).Target = 0
    MapNPC(MapNum, Index).HP = 0
    MapNPC(MapNum, Index).MP = 0
    MapNPC(MapNum, Index).SP = 0
    MapNPC(MapNum, Index).X = 0
    MapNPC(MapNum, Index).Y = 0
    MapNPC(MapNum, Index).Dir = 0

    ' Server use only
    MapNPC(MapNum, Index).SpawnWait = 0
    MapNPC(MapNum, Index).AttackTimer = 0
End Sub

Sub ClearMapNpcs()
    Dim X As Long
    Dim Y As Long

    For Y = 1 To MAX_MAPS
        For X = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(X, Y)
        Next X
    Next Y
End Sub

Sub ClearMap(ByVal MapNum As Long)
    Dim X As Long
    Dim Y As Long

    Map(MapNum).Name = vbNullString
    Map(MapNum).Revision = 0
    Map(MapNum).Moral = 0
    Map(MapNum).Up = 0
    Map(MapNum).Down = 0
    Map(MapNum).Left = 0
    Map(MapNum).Right = 0
    Map(MapNum).Indoors = 0
    Map(MapNum).Weather = 0

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).NPC(X) = 0
    Next X

    For Y = 0 To MAX_MAPY
        For X = 0 To MAX_MAPX
            Map(MapNum).Tile(X, Y).Ground = 0
            Map(MapNum).Tile(X, Y).Mask = 0
            Map(MapNum).Tile(X, Y).Anim = 0
            Map(MapNum).Tile(X, Y).Mask2 = 0
            Map(MapNum).Tile(X, Y).M2Anim = 0
            Map(MapNum).Tile(X, Y).Fringe = 0
            Map(MapNum).Tile(X, Y).FAnim = 0
            Map(MapNum).Tile(X, Y).Fringe2 = 0
            Map(MapNum).Tile(X, Y).F2Anim = 0
            Map(MapNum).Tile(X, Y).Type = 0
            Map(MapNum).Tile(X, Y).Data1 = 0
            Map(MapNum).Tile(X, Y).Data2 = 0
            Map(MapNum).Tile(X, Y).Data3 = 0
            Map(MapNum).Tile(X, Y).String1 = vbNullString
            Map(MapNum).Tile(X, Y).String2 = vbNullString
            Map(MapNum).Tile(X, Y).String3 = vbNullString
            Map(MapNum).Tile(X, Y).Light = 0
            Map(MapNum).Tile(X, Y).GroundSet = 0
            Map(MapNum).Tile(X, Y).MaskSet = 0
            Map(MapNum).Tile(X, Y).AnimSet = 0
            Map(MapNum).Tile(X, Y).Mask2Set = 0
            Map(MapNum).Tile(X, Y).M2AnimSet = 0
            Map(MapNum).Tile(X, Y).FringeSet = 0
            Map(MapNum).Tile(X, Y).FAnimSet = 0
            Map(MapNum).Tile(X, Y).Fringe2Set = 0
            Map(MapNum).Tile(X, Y).F2AnimSet = 0
        Next X
    Next Y

    ' Reset the values for if a player is on the map or not
    PlayersOnMap(MapNum) = NO

    ' Reset the map cache array for this map.
    MapCache(MapNum) = vbNullString
End Sub

Sub ClearMaps()
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call ClearMap(I)
    Next I
End Sub

Sub ClearShop(ByVal Index As Long)
    Dim I As Long

    Shop(Index).Name = vbNullString
    Shop(Index).CurrencyItem = 1
    Shop(Index).FixesItems = 0
    Shop(Index).ShowInfo = 0
    For I = 1 To MAX_SHOP_ITEMS
        Shop(Index).ShopItem(I).ItemNum = 0
        Shop(Index).ShopItem(I).Amount = 0
        Shop(Index).ShopItem(I).Price = 0
    Next I

End Sub

Sub ClearShops()
    Dim I As Long

    For I = 1 To MAX_SHOPS
        Call ClearShop(I)
    Next I
End Sub

Sub ClearSpell(ByVal Index As Long)
    Spell(Index).Name = vbNullString
    Spell(Index).ClassReq = 0
    Spell(Index).LevelReq = 0
    Spell(Index).Type = 0
    Spell(Index).Data1 = 0
    Spell(Index).Data2 = 0
    Spell(Index).Data3 = 0
    Spell(Index).MPCost = 0
    Spell(Index).Sound = 0
    Spell(Index).Range = 0

    Spell(Index).SpellAnim = 0
    Spell(Index).SpellTime = 40
    Spell(Index).SpellDone = 1

    Spell(Index).AE = 0
    Spell(Index).Big = 0

    Spell(Index).Element = 0
End Sub

Sub ClearSpells()
    Dim I As Long

    For I = 1 To MAX_SPELLS
        Call ClearSpell(I)
    Next I
End Sub

Function GetPlayerLogin(ByVal Index As Long) As String
    GetPlayerLogin = Trim$(Player(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Player(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String
    GetPlayerPassword = Trim$(Player(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Player(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String
    GetPlayerName = Trim$(Player(Index).Char(Player(Index).CharNum).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Player(Index).Char(Player(Index).CharNum).Name = Name
End Sub

Function GetPlayerGuild(ByVal Index As Long) As String
    GetPlayerGuild = Trim$(Player(Index).Char(Player(Index).CharNum).Guild)
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal Guild As String)
    Player(Index).Char(Player(Index).CharNum).Guild = Guild
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Long
    GetPlayerGuildAccess = Player(Index).Char(Player(Index).CharNum).GuildAccess
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal GuildAccess As Long)
    Player(Index).Char(Player(Index).CharNum).GuildAccess = GuildAccess
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long
    GetPlayerClass = Player(Index).Char(Player(Index).CharNum).Class
End Function

Sub SetPlayerClassData(ByVal Index As Long, ByVal ClassNum As Long)
    Player(Index).Char(Player(Index).CharNum).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long
    GetPlayerSprite = Player(Index).Char(Player(Index).CharNum).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    If Index > 0 And Index <= MAX_PLAYERS Then
        Player(Index).Char(Player(Index).CharNum).Sprite = Sprite
    End If
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long
    GetPlayerLevel = Player(Index).Char(Player(Index).CharNum).LEVEL
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal LEVEL As Long)
    Player(Index).Char(Player(Index).CharNum).LEVEL = LEVEL
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = Experience(GetPlayerLevel(Index))
End Function

Function GetPlayerExp(ByVal Index As Long) As Long
    GetPlayerExp = Player(Index).Char(Player(Index).CharNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    If GetPlayerLevel(Index) < MAX_LEVEL Then
        Player(Index).Char(Player(Index).CharNum).Exp = Exp
    Else
        Call SetPlayerLevel(Index, MAX_LEVEL)
    End If
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Long
    GetPlayerAccess = Player(Index).Char(Player(Index).CharNum).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Long)
    Player(Index).Char(Player(Index).CharNum).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Long
    GetPlayerPK = Player(Index).Char(Player(Index).CharNum).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Long)
    Player(Index).Char(Player(Index).CharNum).PK = PK
End Sub

Function GetPlayerHP(ByVal Index As Long) As Long
    GetPlayerHP = Player(Index).Char(Player(Index).CharNum).HP
End Function

Sub SetPlayerHP(ByVal Index As Long, ByVal HP As Long)
    Player(Index).Char(Player(Index).CharNum).HP = HP

    If GetPlayerHP(Index) > GetPlayerMaxHP(Index) Then
        Player(Index).Char(Player(Index).CharNum).HP = GetPlayerMaxHP(Index)
    End If
    If GetPlayerHP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).HP = 0
    End If
    'Call SendStats(Index)
End Sub

Function GetPlayerMP(ByVal Index As Long) As Long
    GetPlayerMP = Player(Index).Char(Player(Index).CharNum).MP
End Function

Sub SetPlayerMP(ByVal Index As Long, ByVal MP As Long)
    Player(Index).Char(Player(Index).CharNum).MP = MP

    If GetPlayerMP(Index) > GetPlayerMaxMP(Index) Then
        Player(Index).Char(Player(Index).CharNum).MP = GetPlayerMaxMP(Index)
    End If
    If GetPlayerMP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).MP = 0
    End If
End Sub

Function GetPlayerSP(ByVal Index As Long) As Long
    GetPlayerSP = Player(Index).Char(Player(Index).CharNum).SP
End Function

Sub SetPlayerSP(ByVal Index As Long, ByVal SP As Long)
    Player(Index).Char(Player(Index).CharNum).SP = SP

    If GetPlayerSP(Index) > GetPlayerMaxSP(Index) Then
        Player(Index).Char(Player(Index).CharNum).SP = GetPlayerMaxSP(Index)
    End If
    If GetPlayerSP(Index) < 0 Then
        Player(Index).Char(Player(Index).CharNum).SP = 0
    End If
End Sub

Function GetPlayerMaxHP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).addHP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).addHP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).addHP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).addHP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).addHP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).addHP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).addHP
    End If

    CharNum = Player(Index).CharNum
    ' GetPlayerMaxHP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSTR(index) / 2) + ClassData(Player(index).Char(CharNum).Class).STR) * 2) + add
    GetPlayerMaxHP = (GetPlayerLevel(Index) * addHP.LEVEL) + (GetPlayerSTR(Index) * addHP.STR) + (GetPlayerDEF(Index) * addHP.DEF) + (GetPlayerMAGI(Index) * addHP.Magi) + (GetPlayerSPEED(Index) * addHP.Speed) + Add
End Function

Function GetPlayerMaxMP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).addMP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).addMP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).addMP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).addMP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).addMP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).addMP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).addMP
    End If

    CharNum = Player(Index).CharNum
    ' GetPlayerMaxMP = ((Player(index).Char(CharNum).Level + Int(GetPlayerMAGI(index) / 2) + ClassData(Player(index).Char(CharNum).Class).MAGI) * 2) + add
    GetPlayerMaxMP = (GetPlayerLevel(Index) * addMP.LEVEL) + (GetPlayerSTR(Index) * addMP.STR) + (GetPlayerDEF(Index) * addMP.DEF) + (GetPlayerMAGI(Index) * addMP.Magi) + (GetPlayerSPEED(Index) * addMP.Speed) + Add
End Function

Function GetPlayerMaxSP(ByVal Index As Long) As Long
    Dim CharNum As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).addSP
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).addSP
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).addSP
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).addSP
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).addSP
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).addSP
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).addSP
    End If

    CharNum = Player(Index).CharNum
    ' GetPlayerMaxSP = ((Player(index).Char(CharNum).Level + Int(GetPlayerSPEED(index) / 2) + ClassData(Player(index).Char(CharNum).Class).SPEED) * 2) + add
    GetPlayerMaxSP = (GetPlayerLevel(Index) * addSP.LEVEL) + (GetPlayerSTR(Index) * addSP.STR) + (GetPlayerDEF(Index) * addSP.DEF) + (GetPlayerMAGI(Index) * addSP.Magi) + (GetPlayerSPEED(Index) * addSP.Speed) + Add
End Function

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(ClassData(ClassNum).Name)
End Function

Function GetClassMaxHP(ByVal ClassNum As Long) As Long
    GetClassMaxHP = (1 + Int(ClassData(ClassNum).STR / 2) + ClassData(ClassNum).STR) * 2
End Function

Function GetClassMaxMP(ByVal ClassNum As Long) As Long
    GetClassMaxMP = (1 + Int(ClassData(ClassNum).Magi / 2) + ClassData(ClassNum).Magi) * 2
End Function

Function GetClassMaxSP(ByVal ClassNum As Long) As Long
    GetClassMaxSP = (1 + Int(ClassData(ClassNum).Speed / 2) + ClassData(ClassNum).Speed) * 2
End Function

Function GetClassSTR(ByVal ClassNum As Long) As Long
    GetClassSTR = ClassData(ClassNum).STR
End Function

Function GetClassDEF(ByVal ClassNum As Long) As Long
    GetClassDEF = ClassData(ClassNum).DEF
End Function

Function GetClassSPEED(ByVal ClassNum As Long) As Long
    GetClassSPEED = ClassData(ClassNum).Speed
End Function

Function GetClassMAGI(ByVal ClassNum As Long) As Long
    GetClassMAGI = ClassData(ClassNum).Magi
End Function

Function GetPlayerSTR(ByVal Index As Long) As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddStr
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddStr
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddStr
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddStr
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddStr
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).AddStr
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).AddStr
    End If
    GetPlayerSTR = Player(Index).Char(Player(Index).CharNum).STR + Add
End Function

Sub SetPlayerSTR(ByVal Index As Long, ByVal STR As Long)
    Player(Index).Char(Player(Index).CharNum).STR = STR
End Sub

Function GetPlayerDEF(ByVal Index As Long) As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddDef
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddDef
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddDef
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddDef
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddDef
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).AddDef
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).AddDef
    End If
    GetPlayerDEF = Player(Index).Char(Player(Index).CharNum).DEF + Add
End Function

Sub SetPlayerDEF(ByVal Index As Long, ByVal DEF As Long)
    Player(Index).Char(Player(Index).CharNum).DEF = DEF
End Sub

Function GetPlayerSPEED(ByVal Index As Long) As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddSpeed
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddSpeed
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddSpeed
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddSpeed
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddSpeed
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).AddSpeed
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).AddSpeed
    End If
    GetPlayerSPEED = Player(Index).Char(Player(Index).CharNum).Speed + Add
End Function

Sub SetPlayerSPEED(ByVal Index As Long, ByVal Speed As Long)
    Player(Index).Char(Player(Index).CharNum).Speed = Speed
End Sub

Function GetPlayerMAGI(ByVal Index As Long) As Long
    Dim Add As Long
    Add = 0
    If GetPlayerWeaponSlot(Index) > 0 Then
        Add = Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).AddMagi
    End If
    If GetPlayerArmorSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerArmorSlot(Index))).AddMagi
    End If
    If GetPlayerShieldSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerShieldSlot(Index))).AddMagi
    End If
    If GetPlayerHelmetSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerHelmetSlot(Index))).AddMagi
    End If
    If GetPlayerLegsSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerLegsSlot(Index))).AddMagi
    End If
    If GetPlayerRingSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerRingSlot(Index))).AddMagi
    End If
    If GetPlayerNecklaceSlot(Index) > 0 Then
        Add = Add + Item(GetPlayerInvItemNum(Index, GetPlayerNecklaceSlot(Index))).AddMagi
    End If
    GetPlayerMAGI = Player(Index).Char(Player(Index).CharNum).Magi + Add
End Function

Sub SetPlayerMAGI(ByVal Index As Long, ByVal Magi As Long)
    Player(Index).Char(Player(Index).CharNum).Magi = Magi
End Sub

Function GetPlayerPOINTS(ByVal Index As Long) As Long
    GetPlayerPOINTS = Player(Index).Char(Player(Index).CharNum).POINTS
End Function

Sub SetPlayerPOINTS(ByVal Index As Long, ByVal POINTS As Long)
    Player(Index).Char(Player(Index).CharNum).POINTS = POINTS
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerMap = Player(Index).Char(Player(Index).CharNum).Map
    End If
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Long)
    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Player(Index).Char(Player(Index).CharNum).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long
    GetPlayerX = Player(Index).Char(Player(Index).CharNum).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    Player(Index).Char(Player(Index).CharNum).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long
    GetPlayerY = Player(Index).Char(Player(Index).CharNum).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Player(Index).Char(Player(Index).CharNum).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    GetPlayerDir = Player(Index).Char(Player(Index).CharNum).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Long)
    Player(Index).Char(Player(Index).CharNum).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String
    GetPlayerIP = GameServer.Sockets.Item(Index).RemoteAddress
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long) As Long
    If InvSlot > 0 Then
        GetPlayerInvItemNum = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num
    End If
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemValue = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long) As Long
    GetPlayerInvItemDur = Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Inv(InvSlot).Dur = ItemDur
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long) As Long
    GetPlayerSpell = Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Long, ByVal SpellNum As Long)
    Player(Index).Char(Player(Index).CharNum).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerArmorSlot(ByVal Index As Long) As Long
    GetPlayerArmorSlot = Player(Index).Char(Player(Index).CharNum).ArmorSlot
End Function

Sub SetPlayerArmorSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ArmorSlot = InvNum
End Sub

Function GetPlayerWeaponSlot(ByVal Index As Long) As Long
    GetPlayerWeaponSlot = Player(Index).Char(Player(Index).CharNum).WeaponSlot
End Function

Sub SetPlayerWeaponSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).WeaponSlot = InvNum
End Sub

Function GetPlayerHelmetSlot(ByVal Index As Long) As Long
    GetPlayerHelmetSlot = Player(Index).Char(Player(Index).CharNum).HelmetSlot
End Function

Sub SetPlayerHelmetSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).HelmetSlot = InvNum
End Sub

Function GetPlayerShieldSlot(ByVal Index As Long) As Long
    GetPlayerShieldSlot = Player(Index).Char(Player(Index).CharNum).ShieldSlot
End Function

Sub SetPlayerShieldSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).ShieldSlot = InvNum
End Sub
Function GetPlayerLegsSlot(ByVal Index As Long) As Long
    GetPlayerLegsSlot = Player(Index).Char(Player(Index).CharNum).LegsSlot
End Function

Sub SetPlayerLegsSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).LegsSlot = InvNum
End Sub
Function GetPlayerRingSlot(ByVal Index As Long) As Long
    GetPlayerRingSlot = Player(Index).Char(Player(Index).CharNum).RingSlot
End Function

Sub SetPlayerRingSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).RingSlot = InvNum
End Sub
Function GetPlayerNecklaceSlot(ByVal Index As Long) As Long
    GetPlayerNecklaceSlot = Player(Index).Char(Player(Index).CharNum).NecklaceSlot
End Function

Sub SetPlayerNecklaceSlot(ByVal Index As Long, InvNum As Long)
    Player(Index).Char(Player(Index).CharNum).NecklaceSlot = InvNum
End Sub

Sub BattleMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Byte, ByVal Side As Byte)
    Call SendDataTo(Index, "damagedisplay" & SEP_CHAR & Side & SEP_CHAR & Msg & SEP_CHAR & Color & END_CHAR)
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemNum = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemNum As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).num = ItemNum
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemValue = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemValue As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Value = ItemValue
    Call SendBankUpdate(Index, BankSlot)
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long) As Long
    GetPlayerBankItemDur = Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Dur
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Long, ByVal ItemDur As Long)
    Player(Index).Char(Player(Index).CharNum).Bank(BankSlot).Dur = ItemDur
End Sub

Function GetPlayerHead(ByVal Index As Long) As Integer
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerHead = Player(Index).Char(Player(Index).CharNum).Head
    End If
End Function

Sub SetPlayerHead(ByVal Index As Long, ByVal Head As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).Char(Player(Index).CharNum).Head = Head
    End If
End Sub

Function GetPlayerBody(ByVal Index As Long) As Integer
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerBody = Player(Index).Char(Player(Index).CharNum).Body
    End If
End Function

Sub SetPlayerBody(ByVal Index As Long, ByVal Body As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).Char(Player(Index).CharNum).Body = Body
    End If
End Sub

Function GetPlayerleg(ByVal Index As Long) As Integer
    If Index > 0 And Index < MAX_PLAYERS Then
        GetPlayerleg = Player(Index).Char(Player(Index).CharNum).Leg
    End If
End Function

Sub SetPlayerLeg(ByVal Index As Long, ByVal Leg As Long)
    If Index > 0 And Index < MAX_PLAYERS Then
        Player(Index).Char(Player(Index).CharNum).Leg = Leg
    End If
End Sub

Function GetPlayerPaperdoll(ByVal Index As Long) As Byte
    If Index < MAX_PLAYERS And Index > 0 Then
        If Player(Index).InGame Then
            GetPlayerPaperdoll = Player(Index).Char(Player(Index).CharNum).PAPERDOLL
        End If
    End If
End Function

Sub SetPlayerPaperdoll(ByVal Index As Long, ByVal Mode As Byte)
    If Index < MAX_PLAYERS And Index > 0 Then
        If Mode = 0 Or Mode = 1 Then
            If Player(Index).InGame Then
                Player(Index).Char(Player(Index).CharNum).PAPERDOLL = Mode
            End If
        End If
    End If
End Sub

Function GetSpellReqLevel(ByVal SpellNum As Long) As Long
    GetSpellReqLevel = Spell(SpellNum).LevelReq
End Function

Function GetPlayerTargetNpc(ByVal Index As Long) As Long
    GetPlayerTargetNpc = Player(Index).TargetNPC
End Function
