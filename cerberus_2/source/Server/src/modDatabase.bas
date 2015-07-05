Attribute VB_Name = "modDatabase"
'   This file is part of the Cerberus Engine 2nd Edition.
'
'    The Cerberus Engine 2nd Edition is free software; you can redistribute it
'    and/or modify it under the terms of the GNU General Public License as
'    published by the Free Software Foundation; either version 2 of the License,
'    or (at your option) any later version.
'
'    Cerberus 2nd Edition is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with Cerberus 2nd Edition; if not, write to the Free Software
'    Foundation, Inc., 51 Franklin St, Fifth Floor, Boston, MA  02110-1301  USA

Option Explicit

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = ""
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal FileName As String, Optional RAW As Boolean = False) As Boolean
    If RAW = False Then
        If Dir(App.Path & "\" & FileName) = "" Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
            Exit Function
        End If
    Else
        If Dir(FileName) = "" Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
        End If
    End If
End Function

Sub AddLog(ByVal Text As String, ByVal FN As String)
Dim FileName As String
Dim CheckFN As String
Dim f As Long

    If ServerLog = True Then
        FileName = App.Path & "\logs\" & FN
        CheckFN = "logs\" & FN
    
        If Not FileExist(CheckFN) Then
            f = FreeFile
            Open FileName For Output As #f
            Close #f
        End If
        
        f = FreeFile
        Open FileName For Append As #f
            Print #f, Date & " " & Time & ": " & Text
        Close #f
    End If
End Sub

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim FileName As String
Dim RightPassword As String

    PasswordOK = False
    
    If AccountExist(Name) Then
        FileName = App.Path & "\Accounts\" & Trim(Name) & ".ini"
        RightPassword = GetVar(FileName, "GENERAL", "Password")
        
        If UCase(Trim(Password)) = UCase(Trim(RightPassword)) Then
            PasswordOK = True
        End If
    End If
End Function


' *****************
' ** Player data **
' *****************

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim FileName As String
Dim i As Long
Dim n As Long

    Call ClearPlayer(Index)
    
    FileName = App.Path & "\Accounts\" & Trim(Name) & ".ini"

    Player(Index).Login = GetVar(FileName, "GENERAL", "Login")
    Player(Index).Password = GetVar(FileName, "GENERAL", "Password")

    For i = 1 To MAX_CHARS
        ' General
        Player(Index).Char(i).Name = GetVar(FileName, "CHAR" & i, "Name")
        Player(Index).Char(i).Sex = Val(GetVar(FileName, "CHAR" & i, "Sex"))
        Player(Index).Char(i).Class = Val(GetVar(FileName, "CHAR" & i, "Class"))
        Player(Index).Char(i).Sprite = Val(GetVar(FileName, "CHAR" & i, "Sprite"))
        Player(Index).Char(i).Level = Val(GetVar(FileName, "CHAR" & i, "Level"))
        Player(Index).Char(i).EXP = Val(GetVar(FileName, "CHAR" & i, "Exp"))
        Player(Index).Char(i).Access = Val(GetVar(FileName, "CHAR" & i, "Access"))
        Player(Index).Char(i).PK = Val(GetVar(FileName, "CHAR" & i, "PK"))
        Player(Index).Char(i).Guild = Val(GetVar(FileName, "CHAR" & i, "Guild"))
        
        ' Vitals
        Player(Index).Char(i).HP = Val(GetVar(FileName, "CHAR" & i, "HP"))
        Player(Index).Char(i).MP = Val(GetVar(FileName, "CHAR" & i, "MP"))
        Player(Index).Char(i).SP = Val(GetVar(FileName, "CHAR" & i, "SP"))
        
        ' Stats
        Player(Index).Char(i).STR = Val(GetVar(FileName, "CHAR" & i, "STR"))
        Player(Index).Char(i).DEF = Val(GetVar(FileName, "CHAR" & i, "DEF"))
        Player(Index).Char(i).SPEED = Val(GetVar(FileName, "CHAR" & i, "SPEED"))
        Player(Index).Char(i).MAGI = Val(GetVar(FileName, "CHAR" & i, "MAGI"))
        Player(Index).Char(i).DEX = Val(GetVar(FileName, "CHAR" & i, "DEX"))
        Player(Index).Char(i).POINTS = Val(GetVar(FileName, "CHAR" & i, "POINTS"))
        
        ' Worn equipment
        Player(Index).Char(i).WeaponSlot = Val(GetVar(FileName, "CHAR" & i, "WeaponSlot"))
        Player(Index).Char(i).ArmorSlot = Val(GetVar(FileName, "CHAR" & i, "ArmorSlot"))
        Player(Index).Char(i).HelmetSlot = Val(GetVar(FileName, "CHAR" & i, "HelmetSlot"))
        Player(Index).Char(i).ShieldSlot = Val(GetVar(FileName, "CHAR" & i, "ShieldSlot"))
        Player(Index).Char(i).AmuletSlot = Val(GetVar(FileName, "CHAR" & i, "AmuletSlot"))
        Player(Index).Char(i).RingSlot = Val(GetVar(FileName, "CHAR" & i, "RingSlot"))
        Player(Index).Char(i).ArrowSlot = Val(GetVar(FileName, "CHAR" & i, "ArrowSlot"))
        
        ' Position
        Player(Index).Char(i).Map = Val(GetVar(FileName, "CHAR" & i, "Map"))
        Player(Index).Char(i).x = Val(GetVar(FileName, "CHAR" & i, "X"))
        Player(Index).Char(i).y = Val(GetVar(FileName, "CHAR" & i, "Y"))
        Player(Index).Char(i).Dir = Val(GetVar(FileName, "CHAR" & i, "Dir"))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).x = START_X
            Player(Index).Char(i).y = START_Y
        End If
        
        ' Inventory
        For n = 1 To MAX_INV
            Player(Index).Char(i).Inv(n).Num = Val(GetVar(FileName, "CHAR" & i, "InvItemNum" & n))
            Player(Index).Char(i).Inv(n).Value = Val(GetVar(FileName, "CHAR" & i, "InvItemVal" & n))
            Player(Index).Char(i).Inv(n).Dur = Val(GetVar(FileName, "CHAR" & i, "InvItemDur" & n))
        Next n
        
        ' Skills
        For n = 1 To MAX_PLAYER_SKILLS
            Player(Index).Char(i).Skills(n).Num = Val(GetVar(FileName, "CHAR" & i, "SkillNum" & n))
            Player(Index).Char(i).Skills(n).Level = Val(GetVar(FileName, "CHAR" & i, "SkillLevel" & n))
            Player(Index).Char(i).Skills(n).EXP = Val(GetVar(FileName, "CHAR" & i, "SkillEXP" & n))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(i).Spells(n).Num = Val(GetVar(FileName, "CHAR" & i, "SpellNum" & n))
            Player(Index).Char(i).Spells(n).Level = Val(GetVar(FileName, "CHAR" & i, "SpellLevel" & n))
            Player(Index).Char(i).Spells(n).EXP = Val(GetVar(FileName, "CHAR" & i, "SpellEXP" & n))
        Next n
        
        ' Quests
        For n = 1 To MAX_PLAYER_QUESTS
            Player(Index).Char(i).Quests(n).Num = Val(GetVar(FileName, "CHAR" & i, "QuestNum" & n))
            Player(Index).Char(i).Quests(n).SetMap = Val(GetVar(FileName, "CHAR" & i, "SetMap" & n))
            Player(Index).Char(i).Quests(n).SetBy = Val(GetVar(FileName, "CHAR" & i, "SetBy" & n))
            Player(Index).Char(i).Quests(n).Value = Val(GetVar(FileName, "CHAR" & i, "Value" & n))
            Player(Index).Char(i).Quests(n).Count = Val(GetVar(FileName, "CHAR" & i, "Count" & n))
        Next n
        
        ' Player Maps
        For n = 1 To MAX_PLAYER_MAPS
            Player(Index).Char(i).Maps(n).Num = Val(GetVar(FileName, "CHAR" & i, "MapNum" & n))
        Next n
    Next i
End Sub

Sub SavePlayer(ByVal Index As Long)
Dim FileName As String
Dim i As Long
Dim n As Long

    FileName = App.Path & "\Accounts\" & Trim(Player(Index).Login) & ".ini"
    
    Call PutVar(FileName, "GENERAL", "Login", Trim(Player(Index).Login))
    Call PutVar(FileName, "GENERAL", "Password", Trim(Player(Index).Password))

    For i = 1 To MAX_CHARS
        ' General
        Call PutVar(FileName, "CHAR" & i, "Name", Trim(Player(Index).Char(i).Name))
        Call PutVar(FileName, "CHAR" & i, "Class", STR(Player(Index).Char(i).Class))
        Call PutVar(FileName, "CHAR" & i, "Sex", STR(Player(Index).Char(i).Sex))
        Call PutVar(FileName, "CHAR" & i, "Sprite", STR(Player(Index).Char(i).Sprite))
        Call PutVar(FileName, "CHAR" & i, "Level", STR(Player(Index).Char(i).Level))
        Call PutVar(FileName, "CHAR" & i, "Exp", STR(Player(Index).Char(i).EXP))
        Call PutVar(FileName, "CHAR" & i, "Access", STR(Player(Index).Char(i).Access))
        Call PutVar(FileName, "CHAR" & i, "PK", STR(Player(Index).Char(i).PK))
        Call PutVar(FileName, "CHAR" & i, "Guild", STR(Player(Index).Char(i).Guild))
        
        ' Vitals
        Call PutVar(FileName, "CHAR" & i, "HP", STR(Player(Index).Char(i).HP))
        Call PutVar(FileName, "CHAR" & i, "MP", STR(Player(Index).Char(i).MP))
        Call PutVar(FileName, "CHAR" & i, "SP", STR(Player(Index).Char(i).SP))
        
        ' Stats
        Call PutVar(FileName, "CHAR" & i, "STR", STR(Player(Index).Char(i).STR))
        Call PutVar(FileName, "CHAR" & i, "DEF", STR(Player(Index).Char(i).DEF))
        Call PutVar(FileName, "CHAR" & i, "SPEED", STR(Player(Index).Char(i).SPEED))
        Call PutVar(FileName, "CHAR" & i, "MAGI", STR(Player(Index).Char(i).MAGI))
        Call PutVar(FileName, "CHAR" & i, "DEX", STR(Player(Index).Char(i).DEX))
        Call PutVar(FileName, "CHAR" & i, "POINTS", STR(Player(Index).Char(i).POINTS))
        
        ' Worn equipment
        Call PutVar(FileName, "CHAR" & i, "ArmorSlot", STR(Player(Index).Char(i).ArmorSlot))
        Call PutVar(FileName, "CHAR" & i, "WeaponSlot", STR(Player(Index).Char(i).WeaponSlot))
        Call PutVar(FileName, "CHAR" & i, "HelmetSlot", STR(Player(Index).Char(i).HelmetSlot))
        Call PutVar(FileName, "CHAR" & i, "ShieldSlot", STR(Player(Index).Char(i).ShieldSlot))
        Call PutVar(FileName, "CHAR" & i, "AmuletSlot", STR(Player(Index).Char(i).AmuletSlot))
        Call PutVar(FileName, "CHAR" & i, "RingSlot", STR(Player(Index).Char(i).RingSlot))
        Call PutVar(FileName, "CHAR" & i, "ArrowSlot", STR(Player(Index).Char(i).ArrowSlot))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).x = START_X
            Player(Index).Char(i).y = START_Y
        End If
            
        ' Position
        Call PutVar(FileName, "CHAR" & i, "Map", STR(Player(Index).Char(i).Map))
        Call PutVar(FileName, "CHAR" & i, "X", STR(Player(Index).Char(i).x))
        Call PutVar(FileName, "CHAR" & i, "Y", STR(Player(Index).Char(i).y))
        Call PutVar(FileName, "CHAR" & i, "Dir", STR(Player(Index).Char(i).Dir))
        
        ' Inventory
        For n = 1 To MAX_INV
            Call PutVar(FileName, "CHAR" & i, "InvItemNum" & n, STR(Player(Index).Char(i).Inv(n).Num))
            Call PutVar(FileName, "CHAR" & i, "InvItemVal" & n, STR(Player(Index).Char(i).Inv(n).Value))
            Call PutVar(FileName, "CHAR" & i, "InvItemDur" & n, STR(Player(Index).Char(i).Inv(n).Dur))
        Next n
        
        ' Skills
        For n = 1 To MAX_PLAYER_SKILLS
            Call PutVar(FileName, "CHAR" & i, "SkillNum" & n, STR(Player(Index).Char(i).Skills(n).Num))
            Call PutVar(FileName, "CHAR" & i, "SkillLevel" & n, STR(Player(Index).Char(i).Skills(n).Level))
            Call PutVar(FileName, "CHAR" & i, "SkillEXP" & n, STR(Player(Index).Char(i).Skills(n).EXP))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Call PutVar(FileName, "CHAR" & i, "SpellNum" & n, STR(Player(Index).Char(i).Spells(n).Num))
            Call PutVar(FileName, "CHAR" & i, "SpellLevel" & n, STR(Player(Index).Char(i).Spells(n).Level))
            Call PutVar(FileName, "CHAR" & i, "SpellEXP" & n, STR(Player(Index).Char(i).Spells(n).EXP))
        Next n
        
        ' Quests
        For n = 1 To MAX_PLAYER_QUESTS
            Call PutVar(FileName, "CHAR" & i, "QuestNum" & n, STR(Player(Index).Char(i).Quests(n).Num))
            Call PutVar(FileName, "CHAR" & i, "SetMap" & n, STR(Player(Index).Char(i).Quests(n).SetMap))
            Call PutVar(FileName, "CHAR" & i, "SetBy" & n, STR(Player(Index).Char(i).Quests(n).SetBy))
            Call PutVar(FileName, "CHAR" & i, "Value" & n, STR(Player(Index).Char(i).Quests(n).Value))
            Call PutVar(FileName, "CHAR" & i, "Count" & n, STR(Player(Index).Char(i).Quests(n).Count))
        Next n
        
        ' Player Maps
        For n = 1 To MAX_PLAYER_MAPS
            Call PutVar(FileName, "CHAR" & i, "MapNum" & n, STR(Player(Index).Char(i).Maps(n).Num))
        Next n
    Next i
End Sub

Sub SaveAllPlayersOnline()
Dim i As Long

    For i = 1 To HighIndex
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next i
End Sub


' ******************
' ** Account Data **
' ******************

Function AccountExist(ByVal Name As String) As Boolean
Dim FileName As String

    FileName = "accounts\" & Trim(Name) & ".ini"
    
    If FileExist(FileName) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim(Player(Index).Char(CharNum).Name) <> "" Then
        CharExist = True
    Else
        CharExist = False
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

    If Trim(Player(Index).Char(CharNum).Name) = "" Then
        Player(Index).CharNum = CharNum
        
        Player(Index).Char(CharNum).Name = Name
        Player(Index).Char(CharNum).Sex = Sex
        Player(Index).Char(CharNum).Class = ClassNum
        
        If Player(Index).Char(CharNum).Sex = SEX_MALE Then
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        Else
            Player(Index).Char(CharNum).Sprite = Class(ClassNum).Sprite
        End If
        
        Player(Index).Char(CharNum).Level = 1
                    
        Player(Index).Char(CharNum).STR = Class(ClassNum).STR
        Player(Index).Char(CharNum).DEF = Class(ClassNum).DEF
        Player(Index).Char(CharNum).SPEED = Class(ClassNum).SPEED
        Player(Index).Char(CharNum).MAGI = Class(ClassNum).MAGI
        Player(Index).Char(CharNum).DEX = Class(ClassNum).DEX
        
        Player(Index).Char(CharNum).Map = Class(ClassNum).StartMap
        Player(Index).Char(CharNum).x = Class(ClassNum).StartX
        Player(Index).Char(CharNum).y = Class(ClassNum).StartY
            
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
            
            If Trim(LCase(s)) = Trim(LCase(Name)) Then
                FindChar = True
                Close #f
                Exit Function
            End If
        Loop
    Close #f
End Function

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
        If Trim(LCase(s)) <> Trim(LCase(Name)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Call Kill(App.Path & "\accounts\chartemp.txt")
End Sub


' **************
' ** Map data **
' **************

Sub LoadMaps()
Dim FileName As String
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        Call SetStatus("Loading map... " & i & "/" & MAX_MAPS)
    
        FileName = App.Path & "\maps\map" & i & ".dat"
        
        f = FreeFile
        Open FileName For Binary As #f
            Get #f, , Map(i)
        Close #f
    
        DoEvents
    Next i
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

Sub CheckMaps()
Dim FileName As String
Dim x As Long
Dim y As Long
Dim i As Long
Dim n As Long

    Call ClearMaps
        
    For i = 1 To MAX_MAPS
        FileName = "maps\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(FileName) Then
            Call SaveMap(i)
        End If
    Next i
End Sub


' ******************
' ** General data **
' ******************

Sub LoadClasses()
Dim FileName As String
Dim i As Long

    Call CheckClasses
    
    FileName = App.Path & "\data\classes.ini"
    
    Max_Classes = Val(GetVar(FileName, "INIT", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        Call SetStatus("Loading class... " & i & "/" & Max_Classes)
    
        Class(i).Name = GetVar(FileName, "CLASS" & i, "Name")
        Class(i).Sprite = GetVar(FileName, "CLASS" & i, "Sprite")
        Class(i).STR = Val(GetVar(FileName, "CLASS" & i, "STR"))
        Class(i).DEF = Val(GetVar(FileName, "CLASS" & i, "DEF"))
        Class(i).SPEED = Val(GetVar(FileName, "CLASS" & i, "SPEED"))
        Class(i).MAGI = Val(GetVar(FileName, "CLASS" & i, "MAGI"))
        Class(i).DEX = Val(GetVar(FileName, "CLASS" & i, "DEX"))
        Class(i).StartMap = Val(GetVar(FileName, "CLASS" & i, "StartMap"))
        Class(i).StartX = Val(GetVar(FileName, "CLASS" & i, "StartX"))
        Class(i).StartY = Val(GetVar(FileName, "CLASS" & i, "StartY"))
        
        DoEvents
    Next i
End Sub

Sub SaveClasses()
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\classes.ini"
    
    For i = 0 To Max_Classes
        Call SetStatus("Saving class... " & i & "/" & Max_Classes)
    
        Call PutVar(FileName, "CLASS" & i, "Name", Trim(Class(i).Name))
        Call PutVar(FileName, "CLASS" & i, "Sprite", STR(Class(i).Sprite))
        Call PutVar(FileName, "CLASS" & i, "STR", STR(Class(i).STR))
        Call PutVar(FileName, "CLASS" & i, "DEF", STR(Class(i).DEF))
        Call PutVar(FileName, "CLASS" & i, "SPEED", STR(Class(i).SPEED))
        Call PutVar(FileName, "CLASS" & i, "MAGI", STR(Class(i).MAGI))
        Call PutVar(FileName, "CLASS" & i, "DEX", STR(Class(i).DEX))
        Call PutVar(FileName, "CLASS" & i, "StartMap", STR(Class(i).StartMap))
        Call PutVar(FileName, "CLASS" & i, "StartX", STR(Class(i).StartX))
        Call PutVar(FileName, "CLASS" & i, "StartY", STR(Class(i).StartY))
    Next i
End Sub

Sub CheckClasses()
   If Not FileExist("data\classes.ini") Then
        Call SaveClasses
    End If
End Sub

Sub LoadNpcs()
Dim FileName As String
Dim i As Long
Dim n As Long

    Call CheckNpcs
    
    FileName = App.Path & "\data\npcs.ini"
    
    For i = 1 To MAX_NPCS
        Call SetStatus("Loading npc... " & i & "/" & MAX_NPCS)
    
        Npc(i).Name = GetVar(FileName, "NPC" & i, "Name")
        Npc(i).Sprite = GetVar(FileName, "NPC" & i, "Sprite")
        Npc(i).SpawnSecs = GetVar(FileName, "NPC" & i, "SpawnSecs")
        Npc(i).Behavior = GetVar(FileName, "NPC" & i, "Behavior")
        Npc(i).Range = GetVar(FileName, "NPC" & i, "Range")
        Npc(i).STR = GetVar(FileName, "NPC" & i, "STR")
        Npc(i).DEF = GetVar(FileName, "NPC" & i, "DEF")
        Npc(i).SPEED = GetVar(FileName, "NPC" & i, "SPEED")
        Npc(i).MAGI = GetVar(FileName, "NPC" & i, "MAGI")
        Npc(i).Big = GetVar(FileName, "NPC" & i, "Big")
        Npc(i).MaxHp = GetVar(FileName, "NPC" & i, "MaxHp")
        Npc(i).Respawn = GetVar(FileName, "NPC" & i, "Respawn")
        Npc(i).HitOnlyWith = GetVar(FileName, "NPC" & i, "HitOnlyWith")
        Npc(i).ShopLink = GetVar(FileName, "NPC" & i, "ShopLink")
        Npc(i).ExpType = GetVar(FileName, "NPC" & i, "ExpType")
        Npc(i).EXP = GetVar(FileName, "NPC" & i, "EXP")
        For n = 1 To MAX_NPC_QUESTS
            Npc(i).QuestNPC(n) = GetVar(FileName, "NPC" & i, "Quest" & n)
        Next n
        For n = 1 To MAX_NPC_DROPS
            Npc(i).ItemNPC(n).Chance = GetVar(FileName, "NPC" & i, "Chance" & n)
            Npc(i).ItemNPC(n).ItemNum = GetVar(FileName, "NPC" & i, "ItemNum" & n)
            Npc(i).ItemNPC(n).ItemValue = GetVar(FileName, "NPC" & i, "ItemValue" & n)
        Next n
    
        DoEvents
    Next i
End Sub

Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next i
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\npcs.ini"
    Call SetStatus("Saving npc... " & NpcNum)
    
    Call PutVar(FileName, "NPC" & NpcNum, "Name", Trim(Npc(NpcNum).Name))
    Call PutVar(FileName, "NPC" & NpcNum, "Sprite", Trim(Npc(NpcNum).Sprite))
    Call PutVar(FileName, "NPC" & NpcNum, "SpawnSecs", Trim(Npc(NpcNum).SpawnSecs))
    Call PutVar(FileName, "NPC" & NpcNum, "Behavior", Trim(Npc(NpcNum).Behavior))
    Call PutVar(FileName, "NPC" & NpcNum, "Range", Trim(Npc(NpcNum).Range))
    Call PutVar(FileName, "NPC" & NpcNum, "STR", Trim(Npc(NpcNum).STR))
    Call PutVar(FileName, "NPC" & NpcNum, "DEF", Trim(Npc(NpcNum).DEF))
    Call PutVar(FileName, "NPC" & NpcNum, "SPEED", Trim(Npc(NpcNum).SPEED))
    Call PutVar(FileName, "NPC" & NpcNum, "MAGI", Trim(Npc(NpcNum).MAGI))
    Call PutVar(FileName, "NPC" & NpcNum, "Big", Trim(Npc(NpcNum).Big))
    Call PutVar(FileName, "NPC" & NpcNum, "MaxHp", Trim(Npc(NpcNum).MaxHp))
    Call PutVar(FileName, "NPC" & NpcNum, "Respawn", Trim(Npc(NpcNum).Respawn))
    Call PutVar(FileName, "NPC" & NpcNum, "HitOnlyWith", Trim(Npc(NpcNum).HitOnlyWith))
    Call PutVar(FileName, "NPC" & NpcNum, "ShopLink", Trim(Npc(NpcNum).ShopLink))
    Call PutVar(FileName, "NPC" & NpcNum, "ExpType", Trim(Npc(NpcNum).ExpType))
    Call PutVar(FileName, "NPC" & NpcNum, "EXP", Trim(Npc(NpcNum).EXP))
    For i = 1 To MAX_NPC_QUESTS
        Call PutVar(FileName, "NPC" & NpcNum, "Quest" & i, Trim(Npc(NpcNum).QuestNPC(i)))
    Next i
    For i = 1 To MAX_NPC_DROPS
        Call PutVar(FileName, "NPC" & NpcNum, "Chance" & i, Trim(Npc(NpcNum).ItemNPC(i).Chance))
        Call PutVar(FileName, "NPC" & NpcNum, "ItemNum" & i, Trim(Npc(NpcNum).ItemNPC(i).ItemNum))
        Call PutVar(FileName, "NPC" & NpcNum, "ItemValue" & i, Trim(Npc(NpcNum).ItemNPC(i).ItemValue))
    Next i
End Sub

Sub CheckNpcs()
    If Not FileExist("data\npcs.ini") Then
        Call SaveNpcs
    End If
End Sub

Sub LoadItems()
Dim FileName As String
Dim i As Long

    Call CheckItems
    
    FileName = App.Path & "\data\items.ini"
    
    For i = 1 To MAX_ITEMS
        Call SetStatus("Loading item... " & i & "/" & MAX_ITEMS)
    
        Item(i).Name = GetVar(FileName, "ITEM" & i, "Name")
        Item(i).Pic = Val(GetVar(FileName, "ITEM" & i, "Pic"))
        Item(i).Type = Val(GetVar(FileName, "ITEM" & i, "Type"))
        Item(i).Data1 = Val(GetVar(FileName, "ITEM" & i, "Data1"))
        Item(i).Data2 = Val(GetVar(FileName, "ITEM" & i, "Data2"))
        Item(i).Data3 = Val(GetVar(FileName, "ITEM" & i, "Data3"))
        
        DoEvents
    Next i
End Sub

Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next i
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim FileName As String

    FileName = App.Path & "\data\items.ini"
    Call SetStatus("Saving item... " & ItemNum)
    
    Call PutVar(FileName, "ITEM" & ItemNum, "Name", Trim(Item(ItemNum).Name))
    Call PutVar(FileName, "ITEM" & ItemNum, "Pic", Trim(Item(ItemNum).Pic))
    Call PutVar(FileName, "ITEM" & ItemNum, "Type", Trim(Item(ItemNum).Type))
    Call PutVar(FileName, "ITEM" & ItemNum, "Data1", Trim(Item(ItemNum).Data1))
    Call PutVar(FileName, "ITEM" & ItemNum, "Data2", Trim(Item(ItemNum).Data2))
    Call PutVar(FileName, "ITEM" & ItemNum, "Data3", Trim(Item(ItemNum).Data3))
End Sub

Sub CheckItems()
    If Not FileExist("data\items.ini") Then
        Call SaveItems
    End If
End Sub

Sub LoadShops()
Dim FileName As String
Dim x As Long, y As Long, n As Long

    Call CheckShops
    
    FileName = App.Path & "\data\shops.ini"
    
    For y = 1 To MAX_SHOPS
        Call SetStatus("Loading shops... " & y & "/" & MAX_SHOPS)
    
        Shop(y).Name = GetVar(FileName, "SHOP" & y, "Name")
        Shop(y).FixesItems = GetVar(FileName, "SHOP" & y, "FixesItems")
        
        For x = 1 To MAX_TRADES
            For n = 1 To MAX_GIVE_ITEMS
                Shop(y).TradeItem(x).GiveItem(n) = GetVar(FileName, "SHOP" & y, "Trade" & x & "GiveItem" & n)
                Shop(y).TradeItem(x).GiveValue(n) = GetVar(FileName, "SHOP" & y, "Trade" & x & "GiveValue" & n)
            Next n
            For n = 1 To MAX_GET_ITEMS
                Shop(y).TradeItem(x).GetItem(n) = GetVar(FileName, "SHOP" & y, "Trade" & x & "GetItem" & n)
                Shop(y).TradeItem(x).GetValue(n) = GetVar(FileName, "SHOP" & y, "Trade" & x & "GetValue" & n)
            Next n
            Shop(y).ItemStock(x) = GetVar(FileName, "SHOP" & y, "Trade" & x & "Stock")
        Next x
    
        DoEvents
    Next y
End Sub

Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next i
End Sub

Sub SaveShop(ByVal ShopNum As Long)
Dim FileName As String
Dim i As Long
Dim n As Long

    FileName = App.Path & "\data\shops.ini"
    Call SetStatus("Saving shop... " & ShopNum)
    
    Call PutVar(FileName, "SHOP" & ShopNum, "Name", Trim(Shop(ShopNum).Name))
    Call PutVar(FileName, "SHOP" & ShopNum, "FixesItems", Trim(Shop(ShopNum).FixesItems))
    
    For i = 1 To MAX_TRADES
        For n = 1 To MAX_GIVE_ITEMS
            Call PutVar(FileName, "SHOP" & ShopNum, "Trade" & i & "GiveItem" & n, Trim(Shop(ShopNum).TradeItem(i).GiveItem(n)))
            Call PutVar(FileName, "SHOP" & ShopNum, "Trade" & i & "GiveValue" & n, Trim(Shop(ShopNum).TradeItem(i).GiveValue(n)))
        Next n
        For n = 1 To MAX_GET_ITEMS
            Call PutVar(FileName, "SHOP" & ShopNum, "Trade" & i & "GetItem" & n, Trim(Shop(ShopNum).TradeItem(i).GetItem(n)))
            Call PutVar(FileName, "SHOP" & ShopNum, "Trade" & i & "GetValue" & n, Trim(Shop(ShopNum).TradeItem(i).GetValue(n)))
        Next n
        Call PutVar(FileName, "SHOP" & ShopNum, "Trade" & i & "Stock", Trim(Shop(ShopNum).ItemStock(i)))
    Next i
End Sub

Sub CheckShops()
    If Not FileExist("data\shops.ini") Then
        Call SaveShops
    End If
End Sub

Sub LoadSkills()
Dim FileName As String
Dim i As Long

    Call CheckSkills
    
    FileName = App.Path & "\data\skills.ini"
    
    For i = 1 To MAX_SKILLS
        Call SetStatus("Loading skill... " & i & "/" & MAX_SKILLS)
    
        Skill(i).Name = GetVar(FileName, "SKILL" & i, "Name")
        Skill(i).SkillSprite = Val(GetVar(FileName, "SKILL" & i, "Sprite"))
        Skill(i).ClassReq = Val(GetVar(FileName, "SKILL" & i, "ClassReq"))
        Skill(i).Type = Val(GetVar(FileName, "SKILL" & i, "Type"))
        Skill(i).Data1 = Val(GetVar(FileName, "SKILL" & i, "Data1"))
        Skill(i).Data2 = Val(GetVar(FileName, "SKILL" & i, "Data2"))
        Skill(i).Data3 = Val(GetVar(FileName, "SKILL" & i, "Data3"))
        
        DoEvents
    Next i
End Sub

Sub SaveSkills()
Dim i As Long

    For i = 1 To MAX_SKILLS
        Call SaveSkill(i)
    Next i
End Sub

Sub SaveSkill(ByVal SkillNum As Long)
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\skills.ini"
    Call SetStatus("Saving skill... " & SkillNum)
    
    Call PutVar(FileName, "SKILL" & SkillNum, "Name", Trim(Skill(SkillNum).Name))
    Call PutVar(FileName, "SKILL" & SkillNum, "Sprite", Trim(Skill(SkillNum).SkillSprite))
    Call PutVar(FileName, "SKILL" & SkillNum, "ClassReq", Trim(Skill(SkillNum).ClassReq))
    Call PutVar(FileName, "SKILL" & SkillNum, "Type", Trim(Skill(SkillNum).Type))
    Call PutVar(FileName, "SKILL" & SkillNum, "Data1", Trim(Skill(SkillNum).Data1))
    Call PutVar(FileName, "SKILL" & SkillNum, "Data2", Trim(Skill(SkillNum).Data2))
    Call PutVar(FileName, "SKILL" & SkillNum, "Data3", Trim(Skill(SkillNum).Data3))
End Sub

Sub CheckSkills()
    If Not FileExist("data\skills.ini") Then
        Call SaveSkills
    End If
End Sub

Sub LoadSpells()
Dim FileName As String
Dim i As Long

    Call CheckSpells
    
    FileName = App.Path & "\data\spells.ini"
    
    For i = 1 To MAX_SPELLS
        Call SetStatus("Loading spell... " & i & "/" & MAX_SPELLS)
    
        Spell(i).Name = GetVar(FileName, "SPELL" & i, "Name")
        Spell(i).SpellSprite = Val(GetVar(FileName, "SPELL" & i, "Sprite"))
        Spell(i).ClassReq = Val(GetVar(FileName, "SPELL" & i, "ClassReq"))
        Spell(i).Type = Val(GetVar(FileName, "SPELL" & i, "Type"))
        Spell(i).Data1 = Val(GetVar(FileName, "SPELL" & i, "Data1"))
        Spell(i).Data2 = Val(GetVar(FileName, "SPELL" & i, "Data2"))
        Spell(i).Data3 = Val(GetVar(FileName, "SPELL" & i, "Data3"))
        
        DoEvents
    Next i
End Sub

Sub SaveSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next i
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\spells.ini"
    Call SetStatus("Saving spells... " & SpellNum)
    
    Call PutVar(FileName, "SPELL" & SpellNum, "Name", Trim(Spell(SpellNum).Name))
    Call PutVar(FileName, "SPELL" & SpellNum, "Sprite", Trim(Spell(SpellNum).SpellSprite))
    Call PutVar(FileName, "SPELL" & SpellNum, "ClassReq", Trim(Spell(SpellNum).ClassReq))
    Call PutVar(FileName, "SPELL" & SpellNum, "Type", Trim(Spell(SpellNum).Type))
    Call PutVar(FileName, "SPELL" & SpellNum, "Data1", Trim(Spell(SpellNum).Data1))
    Call PutVar(FileName, "SPELL" & SpellNum, "Data2", Trim(Spell(SpellNum).Data2))
    Call PutVar(FileName, "SPELL" & SpellNum, "Data3", Trim(Spell(SpellNum).Data3))
End Sub

Sub CheckSpells()
    If Not FileExist("data\spells.ini") Then
        Call SaveSpells
    End If
End Sub

Sub LoadQuests()
Dim FileName As String
Dim i As Long

    Call CheckQuests
    
    FileName = App.Path & "\data\quests.ini"
    
    For i = 1 To MAX_QUESTS
        Call SetStatus("Loading quest... " & i & "/" & MAX_QUESTS)
    
        Quest(i).Name = GetVar(FileName, "QUEST" & i, "Name")
        Quest(i).SetBy = Val(GetVar(FileName, "QUEST" & i, "SetBy"))
        Quest(i).ClassReq = Val(GetVar(FileName, "QUEST" & i, "ClassReq"))
        Quest(i).LevelMin = Val(GetVar(FileName, "QUEST" & i, "LevelMin"))
        Quest(i).LevelMax = Val(GetVar(FileName, "QUEST" & i, "LevelMax"))
        Quest(i).Type = Val(GetVar(FileName, "QUEST" & i, "Type"))
        Quest(i).Reward = Val(GetVar(FileName, "QUEST" & i, "Reward"))
        Quest(i).RewardValue = Val(GetVar(FileName, "QUEST" & i, "RewardValue"))
        Quest(i).Data1 = Val(GetVar(FileName, "QUEST" & i, "Data1"))
        Quest(i).Data2 = Val(GetVar(FileName, "QUEST" & i, "Data2"))
        Quest(i).Data3 = Val(GetVar(FileName, "QUEST" & i, "Data3"))
        Quest(i).Description = GetVar(FileName, "QUEST" & i, "Description")
        
        DoEvents
    Next i
End Sub

Sub SaveQuests()
Dim i As Long

    For i = 1 To MAX_QUESTS
        Call SaveQuest(i)
    Next i
End Sub

Sub SaveQuest(ByVal QuestNum As Long)
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\quests.ini"
    Call SetStatus("Saving quest... " & QuestNum)
    
    Call PutVar(FileName, "QUEST" & QuestNum, "Name", Trim(Quest(QuestNum).Name))
    Call PutVar(FileName, "QUEST" & QuestNum, "SetBy", Trim(Quest(QuestNum).SetBy))
    Call PutVar(FileName, "QUEST" & QuestNum, "ClassReq", Trim(Quest(QuestNum).ClassReq))
    Call PutVar(FileName, "QUEST" & QuestNum, "LevelMin", Trim(Quest(QuestNum).LevelMin))
    Call PutVar(FileName, "QUEST" & QuestNum, "LevelMax", Trim(Quest(QuestNum).LevelMax))
    Call PutVar(FileName, "QUEST" & QuestNum, "Type", Trim(Quest(QuestNum).Type))
    Call PutVar(FileName, "QUEST" & QuestNum, "Reward", Trim(Quest(QuestNum).Reward))
    Call PutVar(FileName, "QUEST" & QuestNum, "RewardValue", Trim(Quest(QuestNum).RewardValue))
    Call PutVar(FileName, "QUEST" & QuestNum, "Data1", Trim(Quest(QuestNum).Data1))
    Call PutVar(FileName, "QUEST" & QuestNum, "Data2", Trim(Quest(QuestNum).Data2))
    Call PutVar(FileName, "QUEST" & QuestNum, "Data3", Trim(Quest(QuestNum).Data3))
    Call PutVar(FileName, "QUEST" & QuestNum, "Description", Trim(Quest(QuestNum).Description))
End Sub

Sub CheckQuests()
    If Not FileExist("data\quests.ini") Then
        Call SaveQuests
    End If
End Sub


' *********************
' ** Menu & GUI data **
' *********************

Sub LoadGUIS()
Dim FileName As String
Dim i As Long
Dim n As Long

    Call CheckGUIS
    
    FileName = App.Path & "\data\gui.ini"
    
    For i = 1 To MAX_GUIS
        Call SetStatus("Loading GUI... " & i & "/" & MAX_GUIS)
    
        GUI(i).Name = GetVar(FileName, "GUI" & i, "Name")
        GUI(i).Designer = GetVar(FileName, "GUI" & i, "Designer")
        GUI(i).Revision = Val(GetVar(FileName, "GUI" & i, "Revision"))
        For n = 1 To 7
            GUI(i).Background(n).Data1 = Val(GetVar(FileName, "GUI" & i, "Back" & n & "Data1"))
            GUI(i).Background(n).Data2 = Val(GetVar(FileName, "GUI" & i, "Back" & n & "Data2"))
            GUI(i).Background(n).Data3 = Val(GetVar(FileName, "GUI" & i, "Back" & n & "Data3"))
            GUI(i).Background(n).Data4 = Val(GetVar(FileName, "GUI" & i, "Back" & n & "Data4"))
            GUI(i).Background(n).Data5 = Val(GetVar(FileName, "GUI" & i, "Back" & n & "Data5"))
        Next n
        For n = 1 To 5
            GUI(i).Menu(n).Data1 = Val(GetVar(FileName, "GUI" & i, "Menu" & n & "Data1"))
            GUI(i).Menu(n).Data2 = Val(GetVar(FileName, "GUI" & i, "Menu" & n & "Data2"))
            GUI(i).Menu(n).Data3 = Val(GetVar(FileName, "GUI" & i, "Menu" & n & "Data3"))
            GUI(i).Menu(n).Data4 = Val(GetVar(FileName, "GUI" & i, "Menu" & n & "Data4"))
        Next n
        For n = 1 To 4
            GUI(i).Login(n).Data1 = Val(GetVar(FileName, "GUI" & i, "Login" & n & "Data1"))
            GUI(i).Login(n).Data2 = Val(GetVar(FileName, "GUI" & i, "Login" & n & "Data2"))
            GUI(i).Login(n).Data3 = Val(GetVar(FileName, "GUI" & i, "Login" & n & "Data3"))
            GUI(i).Login(n).Data4 = Val(GetVar(FileName, "GUI" & i, "Login" & n & "Data4"))
        Next n
        For n = 1 To 4
            GUI(i).NewAcc(n).Data1 = Val(GetVar(FileName, "GUI" & i, "NewAcc" & n & "Data1"))
            GUI(i).NewAcc(n).Data2 = Val(GetVar(FileName, "GUI" & i, "NewAcc" & n & "Data2"))
            GUI(i).NewAcc(n).Data3 = Val(GetVar(FileName, "GUI" & i, "NewAcc" & n & "Data3"))
            GUI(i).NewAcc(n).Data4 = Val(GetVar(FileName, "GUI" & i, "NewAcc" & n & "Data4"))
        Next n
        For n = 1 To 4
            GUI(i).DelAcc(n).Data1 = Val(GetVar(FileName, "GUI" & i, "DelAcc" & n & "Data1"))
            GUI(i).DelAcc(n).Data2 = Val(GetVar(FileName, "GUI" & i, "DelAcc" & n & "Data2"))
            GUI(i).DelAcc(n).Data3 = Val(GetVar(FileName, "GUI" & i, "DelAcc" & n & "Data3"))
            GUI(i).DelAcc(n).Data4 = Val(GetVar(FileName, "GUI" & i, "DelAcc" & n & "Data4"))
        Next n
        For n = 1 To 2
            GUI(i).Credits(n).Data1 = Val(GetVar(FileName, "GUI" & i, "Credits" & n & "Data1"))
            GUI(i).Credits(n).Data2 = Val(GetVar(FileName, "GUI" & i, "Credits" & n & "Data2"))
            GUI(i).Credits(n).Data3 = Val(GetVar(FileName, "GUI" & i, "Credits" & n & "Data3"))
            GUI(i).Credits(n).Data4 = Val(GetVar(FileName, "GUI" & i, "Credits" & n & "Data4"))
        Next n
        For n = 1 To 5
            GUI(i).Chars(n).Data1 = Val(GetVar(FileName, "GUI" & i, "Chars" & n & "Data1"))
            GUI(i).Chars(n).Data2 = Val(GetVar(FileName, "GUI" & i, "Chars" & n & "Data2"))
            GUI(i).Chars(n).Data3 = Val(GetVar(FileName, "GUI" & i, "Chars" & n & "Data3"))
            GUI(i).Chars(n).Data4 = Val(GetVar(FileName, "GUI" & i, "Chars" & n & "Data4"))
        Next n
        For n = 1 To 14
            GUI(i).NewChar(n).Data1 = Val(GetVar(FileName, "GUI" & i, "NewChar" & n & "Data1"))
            GUI(i).NewChar(n).Data2 = Val(GetVar(FileName, "GUI" & i, "NewChar" & n & "Data2"))
            GUI(i).NewChar(n).Data3 = Val(GetVar(FileName, "GUI" & i, "NewChar" & n & "Data3"))
            GUI(i).NewChar(n).Data4 = Val(GetVar(FileName, "GUI" & i, "NewChar" & n & "Data4"))
        Next n
        
        DoEvents
    Next i
End Sub

Sub SaveGUIS()
Dim i As Long

    For i = 1 To MAX_GUIS
        Call SaveGUI(i)
    Next i
End Sub

Sub SaveGUI(ByVal GUINum As Long)
Dim FileName As String
Dim i As Long

    FileName = App.Path & "\data\gui.ini"
    Call SetStatus("Saving GUI... " & GUINum)
    
    Call PutVar(FileName, "GUI" & GUINum, "Name", Trim(GUI(GUINum).Name))
    Call PutVar(FileName, "GUI" & GUINum, "Designer", Trim(GUI(GUINum).Designer))
    Call PutVar(FileName, "GUI" & GUINum, "Revision", Trim(GUI(GUINum).Revision + 1))
    For i = 1 To 7
        Call PutVar(FileName, "GUI" & GUINum, "Back" & i & "Data1", Trim(GUI(GUINum).Background(i).Data1))
        Call PutVar(FileName, "GUI" & GUINum, "Back" & i & "Data2", Trim(GUI(GUINum).Background(i).Data2))
        Call PutVar(FileName, "GUI" & GUINum, "Back" & i & "Data3", Trim(GUI(GUINum).Background(i).Data3))
        Call PutVar(FileName, "GUI" & GUINum, "Back" & i & "Data4", Trim(GUI(GUINum).Background(i).Data4))
        Call PutVar(FileName, "GUI" & GUINum, "Back" & i & "Data5", Trim(GUI(GUINum).Background(i).Data5))
    Next i
    For i = 1 To 5
        Call PutVar(FileName, "GUI" & GUINum, "Menu" & i & "Data1", Trim(GUI(GUINum).Menu(i).Data1))
        Call PutVar(FileName, "GUI" & GUINum, "Menu" & i & "Data2", Trim(GUI(GUINum).Menu(i).Data2))
        Call PutVar(FileName, "GUI" & GUINum, "Menu" & i & "Data3", Trim(GUI(GUINum).Menu(i).Data3))
        Call PutVar(FileName, "GUI" & GUINum, "Menu" & i & "Data4", Trim(GUI(GUINum).Menu(i).Data4))
    Next i
    For i = 1 To 4
        Call PutVar(FileName, "GUI" & GUINum, "Login" & i & "Data1", Trim(GUI(GUINum).Login(i).Data1))
        Call PutVar(FileName, "GUI" & GUINum, "Login" & i & "Data2", Trim(GUI(GUINum).Login(i).Data2))
        Call PutVar(FileName, "GUI" & GUINum, "Login" & i & "Data3", Trim(GUI(GUINum).Login(i).Data3))
        Call PutVar(FileName, "GUI" & GUINum, "Login" & i & "Data4", Trim(GUI(GUINum).Login(i).Data4))
    Next i
    For i = 1 To 4
        Call PutVar(FileName, "GUI" & GUINum, "NewAcc" & i & "Data1", Trim(GUI(GUINum).NewAcc(i).Data1))
        Call PutVar(FileName, "GUI" & GUINum, "NewAcc" & i & "Data2", Trim(GUI(GUINum).NewAcc(i).Data2))
        Call PutVar(FileName, "GUI" & GUINum, "NewAcc" & i & "Data3", Trim(GUI(GUINum).NewAcc(i).Data3))
        Call PutVar(FileName, "GUI" & GUINum, "NewAcc" & i & "Data4", Trim(GUI(GUINum).NewAcc(i).Data4))
    Next i
    For i = 1 To 4
        Call PutVar(FileName, "GUI" & GUINum, "DelAcc" & i & "Data1", Trim(GUI(GUINum).DelAcc(i).Data1))
        Call PutVar(FileName, "GUI" & GUINum, "DelAcc" & i & "Data2", Trim(GUI(GUINum).DelAcc(i).Data2))
        Call PutVar(FileName, "GUI" & GUINum, "DelAcc" & i & "Data3", Trim(GUI(GUINum).DelAcc(i).Data3))
        Call PutVar(FileName, "GUI" & GUINum, "DelAcc" & i & "Data4", Trim(GUI(GUINum).DelAcc(i).Data4))
    Next i
    For i = 1 To 2
        Call PutVar(FileName, "GUI" & GUINum, "Credits" & i & "Data1", Trim(GUI(GUINum).Credits(i).Data1))
        Call PutVar(FileName, "GUI" & GUINum, "Credits" & i & "Data2", Trim(GUI(GUINum).Credits(i).Data2))
        Call PutVar(FileName, "GUI" & GUINum, "Credits" & i & "Data3", Trim(GUI(GUINum).Credits(i).Data3))
        Call PutVar(FileName, "GUI" & GUINum, "Credits" & i & "Data4", Trim(GUI(GUINum).Credits(i).Data4))
    Next i
    For i = 1 To 5
        Call PutVar(FileName, "GUI" & GUINum, "Chars" & i & "Data1", Trim(GUI(GUINum).Chars(i).Data1))
        Call PutVar(FileName, "GUI" & GUINum, "Chars" & i & "Data2", Trim(GUI(GUINum).Chars(i).Data2))
        Call PutVar(FileName, "GUI" & GUINum, "Chars" & i & "Data3", Trim(GUI(GUINum).Chars(i).Data3))
        Call PutVar(FileName, "GUI" & GUINum, "Chars" & i & "Data4", Trim(GUI(GUINum).Chars(i).Data4))
    Next i
    For i = 1 To 14
        Call PutVar(FileName, "GUI" & GUINum, "NewChar" & i & "Data1", Trim(GUI(GUINum).NewChar(i).Data1))
        Call PutVar(FileName, "GUI" & GUINum, "NewChar" & i & "Data2", Trim(GUI(GUINum).NewChar(i).Data2))
        Call PutVar(FileName, "GUI" & GUINum, "NewChar" & i & "Data3", Trim(GUI(GUINum).NewChar(i).Data3))
        Call PutVar(FileName, "GUI" & GUINum, "NewChar" & i & "Data4", Trim(GUI(GUINum).NewChar(i).Data4))
    Next i
End Sub

Sub CheckGUIS()
    If Not FileExist("data\gui.ini") Then
        Call SaveGUIS
    End If
End Sub
