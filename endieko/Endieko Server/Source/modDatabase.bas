Attribute VB_Name = "modDatabase"
Option Explicit

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

Public START_MAP As Long
Public START_X As Long
Public START_Y As Long

Public Const ADMIN_LOG = "admin.txt"
Public Const PLAYER_LOG = "player.txt"

Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space$(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Sub LoadExps()
Dim FileName As String
Dim i As Long

    Call CheckExps
    
    FileName = App.Path & "\experience.ini"
    
    For i = 1 To MAX_LEVEL
        'Call SetStatus("Loading exp... " & i & "/" & MAX_LEVEL)
        Experience(i) = GetVar(FileName, "EXPERIENCE", "Exp" & i)
        
        DoEvents
    Next i
End Sub

Sub CheckExps()
    If Not FileExist("experience.ini") Then
        Dim i As Long
    
        For i = 1 To MAX_LEVEL
            'Call SetStatus("Saving exp... " & i & "/" & MAX_LEVEL)
            DoEvents
            Call PutVar(App.Path & "\experience.ini", "EXPERIENCE", "Exp" & i, i * 1500)
        Next i
    End If
End Sub

Sub ClearExps()
Dim i As Long

    For i = 1 To MAX_LEVEL
        Experience(i) = 0
    Next i
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/16/2005  Shannara   Optimized function.
'****************************************************************
    
    If RAW = False Then
        If Dir$(App.Path & "\" & FileName) = vbNullString Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
            Exit Function
        End If
    Else
        If Dir$(FileName) = vbNullString Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
        End If
    End If
End Function

Sub SavePlayer(ByVal Index As Long)
Dim FileName As String
Dim i As Long
Dim n As Long
Dim nFileNum As Integer
Dim nLen As Integer, lCount As Long

    FileName = App.Path & "\Accounts\" & Trim$(Player(Index).Login) & ".end"
    
    If FileExist(FileName, True) Then Kill FileName
    nFileNum = FreeFile
    Open FileName For Binary As #nFileNum
    

    Put #nFileNum, , Player(Index).Login
    Put #nFileNum, , Player(Index).Password
    
    For i = 1 To MAX_CHARS
        Put #nFileNum, , Player(Index).Char(i).Name
        Put #nFileNum, , Player(Index).Char(i).Class
        Put #nFileNum, , Player(Index).Char(i).Sex
        Put #nFileNum, , Player(Index).Char(i).Sprite
        Put #nFileNum, , Player(Index).Char(i).Level
        Put #nFileNum, , Player(Index).Char(i).Exp
        Put #nFileNum, , Player(Index).Char(i).Access
        Put #nFileNum, , Player(Index).Char(i).PK
        Put #nFileNum, , Player(Index).Char(i).Guild
        
        Put #nFileNum, , Player(Index).Char(i).HP
        Put #nFileNum, , Player(Index).Char(i).MP
        Put #nFileNum, , Player(Index).Char(i).SP
        
        Put #nFileNum, , Player(Index).Char(i).STR
        Put #nFileNum, , Player(Index).Char(i).DEF
        Put #nFileNum, , Player(Index).Char(i).SPEED
        Put #nFileNum, , Player(Index).Char(i).MAGI
        Put #nFileNum, , Player(Index).Char(i).POINTS
        
        Put #nFileNum, , Player(Index).Char(i).ArmorSlot
        Put #nFileNum, , Player(Index).Char(i).WeaponSlot
        Put #nFileNum, , Player(Index).Char(i).HelmetSlot
        Put #nFileNum, , Player(Index).Char(i).ShieldSlot
        Put #nFileNum, , Player(Index).Char(i).LegSlot
        Put #nFileNum, , Player(Index).Char(i).BootSlot
        
        Put #nFileNum, , Player(Index).Char(i).BroadcastMute
        Put #nFileNum, , Player(Index).Char(i).GlobalMute
        Put #nFileNum, , Player(Index).Char(i).AdminMute
        Put #nFileNum, , Player(Index).Char(i).MapMute
        Put #nFileNum, , Player(Index).Char(i).EmotMute
        Put #nFileNum, , Player(Index).Char(i).PrivMute
        Put #nFileNum, , Player(Index).Char(i).GuildMute
        Put #nFileNum, , Player(Index).Char(i).PartyMute
        Put #nFileNum, , Player(Index).Char(i).Jailed
        
        Put #nFileNum, , Player(Index).Char(i).Alignment
        Put #nFileNum, , Player(Index).Char(i).FishingLevel
        Put #nFileNum, , Player(Index).Char(i).FishingExp
        Put #nFileNum, , Player(Index).Char(i).MiningLevel
        Put #nFileNum, , Player(Index).Char(i).MiningExp
        Put #nFileNum, , Player(Index).Char(i).LumberLevel
        Put #nFileNum, , Player(Index).Char(i).LumberExp
        
        Put #nFileNum, , Player(Index).Char(i).Status
        
        If Player(Index).Char(i).Map = 0 Or Player(Index).Char(i).Map > MAX_MAPS Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).x = START_X
            Player(Index).Char(i).y = START_Y
        End If
        
        Put #nFileNum, , Player(Index).Char(i).Map
        Put #nFileNum, , Player(Index).Char(i).x
        Put #nFileNum, , Player(Index).Char(i).y
        Put #nFileNum, , Player(Index).Char(i).Dir
        
        For n = 1 To MAX_INV
            Put #nFileNum, , Player(Index).Char(i).Inv(n).Num
            Put #nFileNum, , Player(Index).Char(i).Inv(n).Value
            Put #nFileNum, , Player(Index).Char(i).Inv(n).Dur
        Next n
        
        For n = 1 To MAX_BANK
            Put #nFileNum, , Player(Index).Char(i).Bank(n).Num
            Put #nFileNum, , Player(Index).Char(i).Bank(n).Dur
            Put #nFileNum, , Player(Index).Char(i).Bank(n).Value
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            Put #nFileNum, , Player(Index).Char(i).Spell(n)
        Next n
    Next i
    
    Close #nFileNum
    
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim FileName As String
Dim i As Long
Dim n As Long
Dim nFileNum As Integer
Dim nLen As Integer, lCount As Long

    Call ClearPlayer(Index)
    
    FileName = App.Path & "\Accounts\" & Trim$(Name) & ".end"
    
    nFileNum = FreeFile
    Open FileName For Binary As #nFileNum
    
    Get #nFileNum, , Player(Index).Login
    Get #nFileNum, , Player(Index).Password
    
    For i = 1 To MAX_CHARS
        Get #nFileNum, , Player(Index).Char(i).Name
        Get #nFileNum, , Player(Index).Char(i).Sex
        Get #nFileNum, , Player(Index).Char(i).Class
        Get #nFileNum, , Player(Index).Char(i).Sprite
        Get #nFileNum, , Player(Index).Char(i).Level
        Get #nFileNum, , Player(Index).Char(i).Exp
        Get #nFileNum, , Player(Index).Char(i).Access
        Get #nFileNum, , Player(Index).Char(i).PK
        Get #nFileNum, , Player(Index).Char(i).Guild
        
        Get #nFileNum, , Player(Index).Char(i).HP
        Get #nFileNum, , Player(Index).Char(i).MP
        Get #nFileNum, , Player(Index).Char(i).SP
        
        Get #nFileNum, , Player(Index).Char(i).STR
        Get #nFileNum, , Player(Index).Char(i).DEF
        Get #nFileNum, , Player(Index).Char(i).SPEED
        Get #nFileNum, , Player(Index).Char(i).MAGI
        Get #nFileNum, , Player(Index).Char(i).POINTS
        
        Get #nFileNum, , Player(Index).Char(i).ArmorSlot
        Get #nFileNum, , Player(Index).Char(i).WeaponSlot
        Get #nFileNum, , Player(Index).Char(i).HelmetSlot
        Get #nFileNum, , Player(Index).Char(i).ShieldSlot
        Get #nFileNum, , Player(Index).Char(i).LegSlot
        Get #nFileNum, , Player(Index).Char(i).BootSlot
        
        Get #nFileNum, , Player(Index).Char(i).BroadcastMute
        Get #nFileNum, , Player(Index).Char(i).GlobalMute
        Get #nFileNum, , Player(Index).Char(i).AdminMute
        Get #nFileNum, , Player(Index).Char(i).MapMute
        Get #nFileNum, , Player(Index).Char(i).EmotMute
        Get #nFileNum, , Player(Index).Char(i).PrivMute
        Get #nFileNum, , Player(Index).Char(i).GuildMute
        Get #nFileNum, , Player(Index).Char(i).PartyMute
        Get #nFileNum, , Player(Index).Char(i).Jailed
        
        Get #nFileNum, , Player(Index).Char(i).Alignment
        Get #nFileNum, , Player(Index).Char(i).FishingLevel
        Get #nFileNum, , Player(Index).Char(i).FishingExp
        Get #nFileNum, , Player(Index).Char(i).MiningLevel
        Get #nFileNum, , Player(Index).Char(i).MiningExp
        Get #nFileNum, , Player(Index).Char(i).LumberLevel
        Get #nFileNum, , Player(Index).Char(i).LumberExp
        
        Get #nFileNum, , Player(Index).Char(i).Status

        Get #nFileNum, , Player(Index).Char(i).Map
        Get #nFileNum, , Player(Index).Char(i).x
        Get #nFileNum, , Player(Index).Char(i).y
        Get #nFileNum, , Player(Index).Char(i).Dir
        
        If Player(Index).Char(i).Map = 0 Or Player(Index).Char(i).Map > MAX_MAPS Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).x = START_X
            Player(Index).Char(i).y = START_Y
        End If
        
        For n = 1 To MAX_INV
            Get #nFileNum, , Player(Index).Char(i).Inv(n).Num
            Get #nFileNum, , Player(Index).Char(i).Inv(n).Dur
            Get #nFileNum, , Player(Index).Char(i).Inv(n).Value
        Next n
        
        For n = 1 To MAX_BANK
            Get #nFileNum, , Player(Index).Char(i).Bank(n).Num
            Get #nFileNum, , Player(Index).Char(i).Bank(n).Dur
            Get #nFileNum, , Player(Index).Char(i).Bank(n).Value
        Next n
        
        For n = 1 To MAX_PLAYER_SPELLS
            Get #nFileNum, , Player(Index).Char(i).Spell(n)
        Next n
        DoEvents
    Next i
    
    Close #nFileNum
End Sub

Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    FileName = "Accounts\" & Trim$(Name) & ".end"
    
    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim$(Player(Index).Char(CharNum).Name) <> vbNullString Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function GuildExist(ByVal Index As Long, ByVal GuildNum As Long) As Boolean
    If Trim$(Guild(GuildNum).Name) <> vbNullString Then
        GuildExist = True
    Else
        GuildExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim FileName As String
Dim RightPassword As String * NAME_LENGTH
Dim temp As String
Dim nLen As Integer
Dim nFileNum As Integer

    PasswordOK = False
    
    If AccountExist(Name) Then
        FileName = App.Path & "\Accounts\" & Trim$(Name) & ".end"
        nFileNum = FreeFile
        Open FileName For Binary As #nFileNum
        
        Get #nFileNum, 20, RightPassword
        
        Close #nFileNum
        
        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If
    
End Function

Sub AddAccount(ByVal Index As Long, ByVal Name As String, ByVal Password As String)
Dim i As Long

    Player(Index).Login = Name
    Player(Index).Password = Password
    
    For i = 1 To MAX_CHARS
        Call ClearChar(Index, i)
    Next i
    
    Call SavePlayer(Index)
End Sub

Sub AddChar(ByVal Index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Byte, ByVal CharNum As Long)
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
        Player(Index).Char(CharNum).SPEED = Class(ClassNum).SPEED
        Player(Index).Char(CharNum).MAGI = Class(ClassNum).MAGI
        
        If Class(ClassNum).Map <= 0 Then Class(ClassNum).Map = 1
        If Class(ClassNum).x < 0 Or Class(ClassNum).x > MAX_MAPX Then Class(ClassNum).x = Int(Class(ClassNum).x / 2)
        If Class(ClassNum).y < 0 Or Class(ClassNum).y > MAX_MAPY Then Class(ClassNum).y = Int(Class(ClassNum).y / 2)
        Player(Index).Char(CharNum).Map = Class(ClassNum).Map
        Player(Index).Char(CharNum).x = Class(ClassNum).x
        Player(Index).Char(CharNum).y = Class(ClassNum).y
            
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

Sub DelChar(ByVal Index As Long, ByVal CharNum As Long)
Dim f1 As Long, f2 As Long
Dim s As String

    Call DeleteName(Player(Index).Char(CharNum).Name)
    Call ClearChar(Index, CharNum)
    Call SavePlayer(Index)
End Sub

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

Function FindGuild(ByVal Name As String) As Boolean
Dim f As Long
Dim s As String

    FindGuild = False
    
    f = FreeFile
    Open App.Path & "\data\guilds\guildlist.txt" For Input As #f
        Do While Not EOF(f)
            Input #f, s
            
            If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
                FindGuild = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

Sub SaveAllPlayersOnline()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next i
End Sub

Sub LoadClasses()
Dim FileName As String
Dim i As Long

    Call CheckClasses
    
    FileName = App.Path & "\classes.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        'Call SetStatus("Loading classes... " & i & "/" & Max_Classes)
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        Class(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).SPEED = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class(i).MAGI = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        Class(i).Map = Val(GetVar(FileName, "CLASS" & i, "MAP"))
        Class(i).x = Val(GetVar(FileName, "CLASS" & i, "X"))
        Class(i).y = Val(GetVar(FileName, "CLASS" & i, "Y"))
        Class(i).Locked = Val(GetVar(FileName, "CLASS" & i, "Locked"))
        
        DoEvents
    Next i
End Sub

Sub SaveClasses()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\classes.ini"
    
    For i = 0 To Max_Classes
        c 'all SetStatus("Saving classes... " & i & "/" & Max_Classes)
        DoEvents
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "MaleSprite", STR$(Class(i).MaleSprite))
        Call PutVar(FileName, "CLASS" & i, "FemaleSprite", STR$(Class(i).FemaleSprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR$(Class(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR$(Class(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR$(Class(i).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR$(Class(i).MAGI))
        Call PutVar(FileName, "CLASS" & i, "MAP", STR$(Class(i).Map))
        Call PutVar(FileName, "CLASS" & i, "X", STR$(Class(i).x))
        Call PutVar(FileName, "CLASS" & i, "Y", STR$(Class(i).y))
        Call PutVar(FileName, "CLASS" & i, "Locked", STR$(Class(i).Locked))
    Next i
End Sub

Sub CheckClasses()
    If Not FileExist("classes.ini") Then
        Call SaveClasses
    End If
End Sub

Sub LoadClasses2()
Dim FileName As String
Dim i As Long

    Call CheckClasses2
    
    FileName = App.Path & "\FirstClassAdvancement.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class2(0 To Max_Classes) As ClassRec
    
    Call ClearClasses2
    
    For i = 0 To Max_Classes
        'Call SetStatus("Loading first class advandcement... " & i & "/" & Max_Classes)
        Class2(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class2(i).AdvanceFrom = Val(GetVar(FileName, "CLASS" & i, "AdvanceFrom"))
        Class2(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class2(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        Class2(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class2(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class2(i).SPEED = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class2(i).MAGI = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        
        DoEvents
    Next i
End Sub

Sub SaveClasses2()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\FirstClassAdvancement.ini"
    
    For i = 0 To Max_Classes
        'Call SetStatus("Saving first class advandcement... " & i & "/" & Max_Classes)
        DoEvents
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class2(i).Name))
        Call PutVar(FileName, "CLASS" & i, "AdvanceFrom", STR$(Class2(i).AdvanceFrom))
        Call PutVar(FileName, "CLASS" & i, "LevelReq", STR$(Class2(i).LevelReq))
        Call PutVar(FileName, "CLASS" & i, "MaleSprite", STR$(Class2(i).MaleSprite))
        Call PutVar(FileName, "CLASS" & i, "FemaleSprite", STR$(Class2(i).FemaleSprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR$(Class2(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR$(Class2(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR$(Class2(i).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR$(Class2(i).MAGI))
    Next i
End Sub

Sub CheckClasses2()
    If Not FileExist("FirstClassAdvancement.ini") Then
        Call SaveClasses2
    End If
End Sub

Sub Loadclasses3()
Dim FileName As String
Dim i As Long

    Call Checkclasses3
    
    FileName = App.Path & "\SecondClassAdvancement.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class3(0 To Max_Classes) As ClassRec
    
    Call ClearClasses3
    
    For i = 0 To Max_Classes
        'Call SetStatus("Loading second class advandcement... " & i & "/" & Max_Classes)
        Class3(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class3(i).AdvanceFrom = Val(GetVar(FileName, "CLASS" & i, "AdvanceFrom"))
        Class3(i).LevelReq = Val(GetVar(FileName, "CLASS" & i, "LevelReq"))
        Class3(i).MaleSprite = GetVar(FileName, "CLASS" & i, "MaleSprite")
        Class3(i).FemaleSprite = GetVar(FileName, "CLASS" & i, "FemaleSprite")
        Class3(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class3(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class3(i).SPEED = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class3(i).MAGI = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        
        DoEvents
    Next i
End Sub

Sub Saveclasses3()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\SecondClassAdvancement.ini"
    
    For i = 0 To Max_Classes
        'Call SetStatus("Saving second class advandcement... " & i & "/" & Max_Classes)
        DoEvents
        Call PutVar(FileName, "CLASS" & i, "Name", Trim$(Class3(i).Name))
        Call PutVar(FileName, "CLASS" & i, "AdvanceFrom", STR$(Class3(i).AdvanceFrom))
        Call PutVar(FileName, "CLASS" & i, "MaleSprite", STR$(Class3(i).MaleSprite))
        Call PutVar(FileName, "CLASS" & i, "FemaleSprite", STR$(Class3(i).FemaleSprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR$(Class3(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR$(Class3(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR$(Class3(i).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR$(Class3(i).MAGI))
    Next i
End Sub

Sub Checkclasses3()
    If Not FileExist("SecondClassAdvancement.ini") Then
        Call Saveclasses3
    End If
End Sub

Sub SaveEffects()
Dim i As Long
    
    For i = 1 To MAX_EFFECTS
        If Not FileExist("data\Effects\Effect" & i & ".dat") Then
            DoEvents
            Call SaveEffect(i)
        End If
    Next i
End Sub

Sub SaveEffect(ByVal EffectNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\data\effects\effect" & EffectNum & ".dat"
    
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Effect(EffectNum)
    Close #f
End Sub

Sub LoadEffects()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckEffects
    
    For i = 1 To MAX_EFFECTS
        
        FileName = App.Path & "\data\Effects\Effect" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Item(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckEffects()
    Call SaveEffects
End Sub

Sub SaveItems()
Dim i As Long
        
    'Call SetStatus("Saving items... ")
    For i = 1 To MAX_ITEMS
        If Not FileExist("data\items\item" & i & ".dat") Then
            'Call SetStatus("Saving items... " & i & "/" & MAX_ITEMS)
            DoEvents
            Call SaveItem(i)
        End If
    Next i
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String
Dim f  As Long
FileName = App.Path & "\data\items\item" & ItemNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Item(ItemNum)
    Close #f
End Sub

Sub LoadItems()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckItems
    
    For i = 1 To MAX_ITEMS
        'Call SetEffect("Loading items... " & i & "/" & MAX_ITEMS)
        
        FileName = App.Path & "\data\Items\Item" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Item(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckItems()
    Call SaveItems
End Sub

Sub SaveShops()
Dim i As Long

    'Call SetStatus("Saving shops... ")
    For i = 1 To MAX_SHOPS
        If Not FileExist("data\shops\shop" & i & ".dat") Then
            'Call SetStatus("Saving shops... " & i & "/" & MAX_SHOPS)
            DoEvents
            Call SaveShop(i)
        End If
    Next i
End Sub

Sub SaveShop(ByVal ShopNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\data\shops\shop" & ShopNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Shop(ShopNum)
    Close #f
End Sub

Sub LoadShops()
Dim FileName As String
Dim i As Long, f As Long

    Call CheckShops
    
    For i = 1 To MAX_SHOPS
        'Call SetStatus("Loading shops... " & i & "/" & MAX_SHOPS)
        FileName = App.Path & "\data\shops\shop" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Shop(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckShops()
    Call SaveShops
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\data\spells\spells" & SpellNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Spell(SpellNum)
    Close #f
End Sub

Sub SaveSpells()
Dim i As Long

    'Call SetStatus("Saving spells... ")
    For i = 1 To MAX_SPELLS
        If Not FileExist("data\spells\spells" & i & ".dat") Then
            'Call SetStatus("Saving spells... " & i & "/" & MAX_SPELLS)
            DoEvents
            Call SaveSpell(i)
        End If
    Next i
End Sub

Sub LoadSpells()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckSpells
    
    For i = 1 To MAX_SPELLS
        'Call SetStatus("Loading spells... " & i & "/" & MAX_SPELLS)
        
        FileName = App.Path & "\data\spells\spells" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Spell(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckSpells()
    Call SaveSpells
End Sub

Sub SaveNpcs()
Dim i As Long

    'Call SetStatus("Saving npcs... ")
    
    For i = 1 To MAX_NPCS
        If Not FileExist("data\npcs\npc" & i & ".dat") Then
            'Call SetStatus("Saving npcs... " & i & "/" & MAX_NPCS)
            DoEvents
            Call SaveNpc(i)
        End If
    Next i
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String
Dim f As Long
FileName = App.Path & "\data\npcs\npc" & NpcNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Npc(NpcNum)
    Close #f
End Sub

Sub LoadNpcs()
Dim FileName As String
Dim i As Long
Dim z As Long
Dim f As Long

    Call CheckNpcs
        
    For i = 1 To MAX_NPCS
        'Call SetStatus("Loading npcs... " & i & "/" & MAX_NPCS)
        FileName = App.Path & "\data\npcs\npc" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Npc(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckNpcs()
    Call SaveNpcs
End Sub

Sub SaveArrows()
Dim i As Long

    'Call SetStatus("Saving Arrows... ")
    
    For i = 1 To MAX_ARROWS
        If Not FileExist("data\arrows\arrow" & i & ".dat") Then
            'Call SetStatus("Saving Arrows... " & i & "/" & MAX_Arrows)
            DoEvents
            Call SaveArrow(i)
        End If
    Next i
End Sub

Sub SaveArrow(ByVal ArrowNum As Long)
Dim FileName As String
Dim f As Long
FileName = App.Path & "\data\arrows\arrow" & ArrowNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Arrows(ArrowNum)
    Close #f
End Sub

Sub LoadArrows()
Dim FileName As String
Dim i As Long
Dim z As Long
Dim f As Long

    Call CheckArrows
        
    For i = 1 To MAX_ARROWS
        'Call SetStatus("Loading npcs... " & i & "/" & MAX_NPCS)
        FileName = App.Path & "\data\arrows\arrow" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Arrows(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckArrows()
    Call SaveArrows
End Sub

Sub SaveEmoticons()
Dim i As Long

    
    For i = 1 To MAX_EMOTICONS
        If Not FileExist("data\emoticons\emoticon" & i & ".dat") Then
            DoEvents
            Call SaveEmoticon(i)
        End If
    Next i
End Sub

Sub SaveEmoticon(ByVal EmoticonNum As Long)
Dim FileName As String
Dim f As Long
FileName = App.Path & "\data\emoticons\emoticon" & EmoticonNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Emoticons(EmoticonNum)
    Close #f
End Sub

Sub LoadEmoticons()
Dim FileName As String
Dim i As Long
Dim z As Long
Dim f As Long

    Call CheckEmoticons
        
    For i = 1 To MAX_EMOTICONS
        'Call SetStatus("Loading npcs... " & i & "/" & MAX_NPCS)
        FileName = App.Path & "\data\emoticons\emoticon" & i & ".dat"
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Emoticons(i)
        Close #f
        
        DoEvents
    Next i
End Sub

Sub CheckEmoticons()
    Call SaveEmoticons
End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim FileName As String
Dim f As Long

    FileName = App.Path & "\data\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open FileName For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SaveMaps()
Dim FileName As String
Dim i As Long
Dim f As Long

    For i = 1 To MAX_MAPS
        'Call SetStatus("Saving maps... " & i & "/" & MAX_MAPS)
        DoEvents
        Call SaveMap(i)
    Next i
End Sub

Sub LoadMaps()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        'Call SetStatus("Loading maps... " & i & "/" & MAX_MAPS)
        FileName = App.Path & "\data\maps\map" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(i)
        Close #f
    
        DoEvents
    Next i
End Sub

Sub CheckMaps()
Dim FileName As String
Dim x As Long
Dim y As Long
Dim i As Long
Dim n As Long

    Call ClearMaps
        
    For i = 1 To MAX_MAPS
        FileName = "data\maps\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            'Call SetStatus("Saving maps... " & i & "/" & MAX_NPCS)
            DoEvents
            
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String)
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
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim FileName As String, IP As String, bnam As String
Dim f As Long, i As Long
Dim BNum As Integer, b As Integer
Dim MAX_BANS As Long

    FileName = App.Path & "\banlist.ini"
   
    IP = GetPlayerIP(BanPlayerIndex)

    b = 1

    For i = 0 To MAX_BANS + 1
    If Ban(i).BannedIP = vbNullString Then
        BNum = i
        Exit For
    End If
   
    If i = MAX_BANS + 1 Then
        BNum = MAX_BANS + 1
        Exit For
    End If
   
    Next i
   
    ' Add there data to a ban slot
    Ban(BNum).BannedIP = IP
    Ban(BNum).BannedChar = GetPlayerName(BanPlayerIndex)
    Ban(BNum).BannedBy = GetPlayerName(BannedByIndex)
    Ban(BNum).BannedHD = GetPlayerHD(BanPlayerIndex)
    Call SaveBan(BNum)
    Call PutVar(FileName, "Total", "Total", GetVar(FileName, "Total", "Total") + 1)
    MAX_BANS = MAX_BANS + 1
    
    'Alert People
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
End Sub

Sub UnBanIndex(ByVal BannedPlayerName As String, ByVal DeBannedByIndex As Long)
Dim FileName As String, IP As String, bnam As String
Dim f As Long, i As Long
Dim b As Integer
Dim MAX_BANS As Integer

    FileName = App.Path & "\banlist.ini"

    For i = 0 To MAX_BANS + 1
        If LCase$(GetVar(FileName, "Ban" & i, "BannedChar")) = LCase$(BannedPlayerName) Then
                 ' Delete there data to a ban slot
                 Ban(i).BannedIP = vbNullString
                 Ban(i).BannedChar = vbNullString
                 Ban(i).BannedBy = vbNullString
                 Ban(i).BannedHD = vbNullString
                 Call SaveBan(i)
                 
                 'Alert People
                 Call GlobalMsg(BannedPlayerName & " has been unbanned from " & GAME_NAME & " by " & GetPlayerName(DeBannedByIndex) & "!", White)
                 Call AddLog(GetPlayerName(DeBannedByIndex) & " has unbanned " & BannedPlayerName & ".", ADMIN_LOG)
             Exit For
        End If
       
        If i = MAX_BANS + 1 Then
             Call PlayerMsg(DeBannedByIndex, "Player is not banned!", White)
        End If
    Next i

End Sub

'Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
'Dim FileName, IP As String
'Dim f As Long, I As Long
'
'    FileName = App.Path & "\banlist.txt"
'
'    ' Make sure the file exists
'    If Not FileExist("banlist.txt") Then
'        f = FreeFile
'        Open FileName For Output As #f
'        Close #f
'    End If
'
'    ' Cut off last portion of ip
'    IP = GetPlayerIP(BanPlayerIndex)
'
'    For I = Len(IP) To 1 Step -1
'        If Mid$(IP, I, 1) = "." Then
'            Exit For
'        End If
'    Next I
'    IP = Mid$(IP, 1, I)
'
'    f = FreeFile
'    Open FileName For Append As #f
'        Print #f, IP & "," & GetPlayerName(BannedByIndex)
'    Close #f
'
'    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
'    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
'    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
'End Sub

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

Sub DeleteGuild(ByVal Name As String)
Dim f1 As Long, f2 As Long
Dim s As String

    Call FileCopy(App.Path & "\data\guilds\guildlist.txt", App.Path & "\data\guilds\guildtemp.txt")
    
    ' Destroy name from guildlist
    f1 = FreeFile
    Open App.Path & "\data\guilds\guildtemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\guilds\guildlist.txt" For Output As #f2
        
    Do While Not EOF(f1)
        Input #f1, s
        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Call Kill(App.Path & "\data\guilds\guildtemp.txt")
End Sub
