Attribute VB_Name = "modDatabase"
Option Explicit

Public Function GetVar(File As String, Header As String, Var As String) As String
Dim sSpaces As String   ' Max string length
Dim szReturn As String  ' Return default value if not found
  
    szReturn = vbNullString
  
    sSpaces = Space(5000)
  
    Call GetPrivateProfileString(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
  
    GetVar = RTrim(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal Filename As String, Optional RAW As Boolean = False) As Boolean
'****************************************************************
'* WHEN        WHO        WHAT
'* ----        ---        ----
'* 07/16/2005  Shannara   Optimized function.
'****************************************************************
    
    If RAW = False Then
        If LenB(Dir$(App.Path & "\" & Filename)) = 0 Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
            Exit Function
        End If
    Else
        If LenB(Dir$(Filename)) = 0 Then
            FileExist = False
            Exit Function
        Else
            FileExist = True
        End If
    End If
End Function

Sub SavePlayer(ByVal Index As Long)
Dim Filename As String
Dim i As Long
Dim n As Long

    Filename = App.Path & "\Accounts\" & Trim(Player(Index).Login) & ".ini"
    
    Call PutVar(Filename, "GENERAL", "Login", Trim(Player(Index).Login))
    Call PutVar(Filename, "GENERAL", "Password", Trim(Player(Index).Password))

    For i = 1 To MAX_CHARS
        ' General
        Call PutVar(Filename, "CHAR" & i, "Name", Trim(Player(Index).Char(i).Name))
        Call PutVar(Filename, "CHAR" & i, "Class", STR(Player(Index).Char(i).Class))
        Call PutVar(Filename, "CHAR" & i, "Sex", STR(Player(Index).Char(i).Sex))
        Call PutVar(Filename, "CHAR" & i, "Sprite", STR(Player(Index).Char(i).Sprite))
        Call PutVar(Filename, "CHAR" & i, "Level", STR(Player(Index).Char(i).Level))
        Call PutVar(Filename, "CHAR" & i, "Exp", STR(Player(Index).Char(i).Exp))
        Call PutVar(Filename, "CHAR" & i, "Access", STR(Player(Index).Char(i).Access))
        Call PutVar(Filename, "CHAR" & i, "PK", STR(Player(Index).Char(i).PK))
        Call PutVar(Filename, "CHAR" & i, "Guild", STR(Player(Index).Char(i).Guild))
        
        ' Vitals
        Call PutVar(Filename, "CHAR" & i, "HP", STR(Player(Index).Char(i).HP))
        Call PutVar(Filename, "CHAR" & i, "MP", STR(Player(Index).Char(i).MP))
        Call PutVar(Filename, "CHAR" & i, "SP", STR(Player(Index).Char(i).SP))
        
        ' Stats
        Call PutVar(Filename, "CHAR" & i, "STR", STR(Player(Index).Char(i).STR))
        Call PutVar(Filename, "CHAR" & i, "DEF", STR(Player(Index).Char(i).DEF))
        Call PutVar(Filename, "CHAR" & i, "SPEED", STR(Player(Index).Char(i).SPEED))
        Call PutVar(Filename, "CHAR" & i, "MAGI", STR(Player(Index).Char(i).MAGI))
        Call PutVar(Filename, "CHAR" & i, "POINTS", STR(Player(Index).Char(i).POINTS))
        
        ' Worn equipment
        Call PutVar(Filename, "CHAR" & i, "ArmorSlot", STR(Player(Index).Char(i).ArmorSlot))
        Call PutVar(Filename, "CHAR" & i, "WeaponSlot", STR(Player(Index).Char(i).WeaponSlot))
        Call PutVar(Filename, "CHAR" & i, "HelmetSlot", STR(Player(Index).Char(i).HelmetSlot))
        Call PutVar(Filename, "CHAR" & i, "ShieldSlot", STR(Player(Index).Char(i).ShieldSlot))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).x = START_X
            Player(Index).Char(i).y = START_Y
        End If
            
        ' Position
        Call PutVar(Filename, "CHAR" & i, "Map", STR(Player(Index).Char(i).Map))
        Call PutVar(Filename, "CHAR" & i, "X", STR(Player(Index).Char(i).x))
        Call PutVar(Filename, "CHAR" & i, "Y", STR(Player(Index).Char(i).y))
        Call PutVar(Filename, "CHAR" & i, "Dir", STR(Player(Index).Char(i).Dir))
        
        ' Inventory
        For n = 1 To MAX_INV
            Call PutVar(Filename, "CHAR" & i, "InvItemNum" & n, STR(Player(Index).Char(i).Inv(n).Num))
            Call PutVar(Filename, "CHAR" & i, "InvItemVal" & n, STR(Player(Index).Char(i).Inv(n).Value))
            Call PutVar(Filename, "CHAR" & i, "InvItemDur" & n, STR(Player(Index).Char(i).Inv(n).Dur))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Call PutVar(Filename, "CHAR" & i, "Spell" & n, STR(Player(Index).Char(i).Spell(n)))
        Next n
    Next i
End Sub

Sub LoadPlayer(ByVal Index As Long, ByVal Name As String)
Dim Filename As String
Dim i As Long
Dim n As Long

    Call ClearPlayer(Index)
    
    Filename = App.Path & "\Accounts\" & Trim(Name) & ".ini"

    Player(Index).Login = GetVar(Filename, "GENERAL", "Login")
    Player(Index).Password = GetVar(Filename, "GENERAL", "Password")

    For i = 1 To MAX_CHARS
        ' General
        Player(Index).Char(i).Name = GetVar(Filename, "CHAR" & i, "Name")
        Player(Index).Char(i).Sex = Val(GetVar(Filename, "CHAR" & i, "Sex"))
        Player(Index).Char(i).Class = Val(GetVar(Filename, "CHAR" & i, "Class"))
        Player(Index).Char(i).Sprite = Val(GetVar(Filename, "CHAR" & i, "Sprite"))
        Player(Index).Char(i).Level = Val(GetVar(Filename, "CHAR" & i, "Level"))
        Player(Index).Char(i).Exp = Val(GetVar(Filename, "CHAR" & i, "Exp"))
        Player(Index).Char(i).Access = Val(GetVar(Filename, "CHAR" & i, "Access"))
        Player(Index).Char(i).PK = Val(GetVar(Filename, "CHAR" & i, "PK"))
        Player(Index).Char(i).Guild = Val(GetVar(Filename, "CHAR" & i, "Guild"))
        
        ' Vitals
        Player(Index).Char(i).HP = Val(GetVar(Filename, "CHAR" & i, "HP"))
        Player(Index).Char(i).MP = Val(GetVar(Filename, "CHAR" & i, "MP"))
        Player(Index).Char(i).SP = Val(GetVar(Filename, "CHAR" & i, "SP"))
        
        ' Stats
        Player(Index).Char(i).STR = Val(GetVar(Filename, "CHAR" & i, "STR"))
        Player(Index).Char(i).DEF = Val(GetVar(Filename, "CHAR" & i, "DEF"))
        Player(Index).Char(i).SPEED = Val(GetVar(Filename, "CHAR" & i, "SPEED"))
        Player(Index).Char(i).MAGI = Val(GetVar(Filename, "CHAR" & i, "MAGI"))
        Player(Index).Char(i).POINTS = Val(GetVar(Filename, "CHAR" & i, "POINTS"))
        
        ' Worn equipment
        Player(Index).Char(i).ArmorSlot = Val(GetVar(Filename, "CHAR" & i, "ArmorSlot"))
        Player(Index).Char(i).WeaponSlot = Val(GetVar(Filename, "CHAR" & i, "WeaponSlot"))
        Player(Index).Char(i).HelmetSlot = Val(GetVar(Filename, "CHAR" & i, "HelmetSlot"))
        Player(Index).Char(i).ShieldSlot = Val(GetVar(Filename, "CHAR" & i, "ShieldSlot"))
        
        ' Position
        Player(Index).Char(i).Map = Val(GetVar(Filename, "CHAR" & i, "Map"))
        Player(Index).Char(i).x = Val(GetVar(Filename, "CHAR" & i, "X"))
        Player(Index).Char(i).y = Val(GetVar(Filename, "CHAR" & i, "Y"))
        Player(Index).Char(i).Dir = Val(GetVar(Filename, "CHAR" & i, "Dir"))
        
        ' Check to make sure that they aren't on map 0, if so reset'm
        If Player(Index).Char(i).Map = 0 Then
            Player(Index).Char(i).Map = START_MAP
            Player(Index).Char(i).x = START_X
            Player(Index).Char(i).y = START_Y
        End If
        
        ' Inventory
        For n = 1 To MAX_INV
            Player(Index).Char(i).Inv(n).Num = Val(GetVar(Filename, "CHAR" & i, "InvItemNum" & n))
            Player(Index).Char(i).Inv(n).Value = Val(GetVar(Filename, "CHAR" & i, "InvItemVal" & n))
            Player(Index).Char(i).Inv(n).Dur = Val(GetVar(Filename, "CHAR" & i, "InvItemDur" & n))
        Next n
        
        ' Spells
        For n = 1 To MAX_PLAYER_SPELLS
            Player(Index).Char(i).Spell(n) = Val(GetVar(Filename, "CHAR" & i, "Spell" & n))
        Next n
    Next i
End Sub

Function AccountExist(ByVal Name As String) As Boolean
Dim Filename As String

    Filename = "Accounts\" & Trim(Name) & ".ini"
    
    If FileExist(Filename) Then
        AccountExist = True
    Else
        AccountExist = False
    End If
End Function

Function CharExist(ByVal Index As Long, ByVal CharNum As Long) As Boolean
    If Trim(Player(Index).Char(CharNum).Name) <> vbNullString Then
        CharExist = True
    Else
        CharExist = False
    End If
End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim Filename As String
Dim RightPassword As String

    PasswordOK = False
    
    If AccountExist(Name) Then
        Filename = App.Path & "\Accounts\" & Trim(Name) & ".ini"
        RightPassword = GetVar(Filename, "GENERAL", "Password")
        
        If UCase(Trim(Password)) = UCase(Trim(RightPassword)) Then
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

    If Len(Trim(Player(Index).Char(CharNum).Name)) = 0 Then
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
        
        Player(Index).Char(CharNum).Map = START_MAP
        Player(Index).Char(CharNum).x = START_X
        Player(Index).Char(CharNum).y = START_Y
            
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
    Open App.Path & "\Accounts\charlist.txt" For Input As #f
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

Sub SaveAllPlayersOnline()
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            Call SavePlayer(i)
        End If
    Next i
End Sub

Sub LoadClasses()
    Dim Filename As String
    Dim i As Long
    
    Filename = App.Path & "\data\classes.ini"

    Call CheckClasses
        
    Max_Classes = Val(GetVar(Filename, "INIT", "MaxClasses"))
    
    ReDim Class(0 To Max_Classes) As ClassRec
    
    Call ClearClasses
    
    For i = 0 To Max_Classes
        If Cancel_Load Then
            DestroyServer
            Exit Sub
        End If
        
        Class(i).Name = GetVar(Filename, "CLASS" & i, "Name")
        Class(i).Sprite = GetVar(Filename, "CLASS" & i, "Sprite")
        Class(i).STR = Val(GetVar(Filename, "CLASS" & i, "STR"))
        Class(i).DEF = Val(GetVar(Filename, "CLASS" & i, "DEF"))
        Class(i).SPEED = Val(GetVar(Filename, "CLASS" & i, "SPEED"))
        Class(i).MAGI = Val(GetVar(Filename, "CLASS" & i, "MAGI"))
        
        SetStatusProgress frmLoad.pBar.Value + 1
        DoEvents
    Next i
End Sub

Sub SaveClasses()
Dim Filename As String
Dim i As Long

    Filename = App.Path & "\data\classes.ini"
    
    For i = 0 To Max_Classes
        Call PutVar(Filename, "CLASS" & i, "Name", Trim(Class(i).Name))
        Call PutVar(Filename, "CLASS" & i, "Sprite", STR(Class(i).Sprite))
        Call PutVar(Filename, "CLASS" & i, "STR", STR(Class(i).STR))
        Call PutVar(Filename, "CLASS" & i, "DEF", STR(Class(i).DEF))
        Call PutVar(Filename, "CLASS" & i, "SPEED", STR(Class(i).SPEED))
        Call PutVar(Filename, "CLASS" & i, "MAGI", STR(Class(i).MAGI))
    Next i
End Sub

Sub CheckClasses()
    If Not FileExist("data\classes.ini") Then
        Call SaveClasses
    End If
End Sub

Sub SaveItems()
Dim i As Long
    
    For i = 1 To MAX_ITEMS
        Call SaveItem(i)
    Next i
End Sub

Sub SaveItem(ByVal ItemNum As Long)
Dim Filename As String

    Filename = App.Path & "\data\items.ini"
    
    Call PutVar(Filename, "ITEM" & ItemNum, "Name", Trim(Item(ItemNum).Name))
    Call PutVar(Filename, "ITEM" & ItemNum, "Pic", Trim(Item(ItemNum).Pic))
    Call PutVar(Filename, "ITEM" & ItemNum, "Type", Trim(Item(ItemNum).Type))
    Call PutVar(Filename, "ITEM" & ItemNum, "Data1", Trim(Item(ItemNum).Data1))
    Call PutVar(Filename, "ITEM" & ItemNum, "Data2", Trim(Item(ItemNum).Data2))
    Call PutVar(Filename, "ITEM" & ItemNum, "Data3", Trim(Item(ItemNum).Data3))
End Sub

Sub LoadItems()
Dim Filename As String
Dim i As Long

    Call CheckItems
    
    Filename = App.Path & "\data\items.ini"
    
    For i = 1 To MAX_ITEMS
        If Cancel_Load Then
            DestroyServer
            Exit Sub
        End If
    
        Item(i).Name = GetVar(Filename, "ITEM" & i, "Name")
        Item(i).Pic = Val(GetVar(Filename, "ITEM" & i, "Pic"))
        Item(i).Type = Val(GetVar(Filename, "ITEM" & i, "Type"))
        Item(i).Data1 = Val(GetVar(Filename, "ITEM" & i, "Data1"))
        Item(i).Data2 = Val(GetVar(Filename, "ITEM" & i, "Data2"))
        Item(i).Data3 = Val(GetVar(Filename, "ITEM" & i, "Data3"))
        
        SetStatusProgress frmLoad.pBar.Value + 1
        DoEvents
    Next i
End Sub

Sub CheckItems()
    If Not FileExist("data\items.ini") Then
        Call SaveItems
    End If
End Sub

Sub SaveShops()
Dim i As Long

    For i = 1 To MAX_SHOPS
        Call SaveShop(i)
    Next i
End Sub

Sub SaveShop(ByVal ShopNum As Long)
Dim Filename As String
Dim i As Long

    Filename = App.Path & "\data\shops.ini"
    
    Call PutVar(Filename, "SHOP" & ShopNum, "Name", Trim(Shop(ShopNum).Name))
    Call PutVar(Filename, "SHOP" & ShopNum, "JoinSay", Trim(Shop(ShopNum).JoinSay))
    Call PutVar(Filename, "SHOP" & ShopNum, "LeaveSay", Trim(Shop(ShopNum).LeaveSay))
    Call PutVar(Filename, "SHOP" & ShopNum, "FixesItems", Trim(Shop(ShopNum).FixesItems))
    
    For i = 1 To MAX_TRADES
        Call PutVar(Filename, "SHOP" & ShopNum, "GiveItem" & i, Trim(Shop(ShopNum).TradeItem(i).GiveItem))
        Call PutVar(Filename, "SHOP" & ShopNum, "GiveValue" & i, Trim(Shop(ShopNum).TradeItem(i).GiveValue))
        Call PutVar(Filename, "SHOP" & ShopNum, "GetItem" & i, Trim(Shop(ShopNum).TradeItem(i).GetItem))
        Call PutVar(Filename, "SHOP" & ShopNum, "GetValue" & i, Trim(Shop(ShopNum).TradeItem(i).GetValue))
    Next i
End Sub

Sub LoadShops()
On Error Resume Next

Dim Filename As String
Dim x As Long, y As Long

    Call CheckShops
    
    Filename = App.Path & "\data\shops.ini"
    
    For y = 1 To MAX_SHOPS
        If Cancel_Load Then
            DestroyServer
            Exit Sub
        End If
    
        Shop(y).Name = GetVar(Filename, "SHOP" & y, "Name")
        Shop(y).JoinSay = GetVar(Filename, "SHOP" & y, "JoinSay")
        Shop(y).LeaveSay = GetVar(Filename, "SHOP" & y, "LeaveSay")
        Shop(y).FixesItems = GetVar(Filename, "SHOP" & y, "FixesItems")
        
        For x = 1 To MAX_TRADES
            Shop(y).TradeItem(x).GiveItem = GetVar(Filename, "SHOP" & y, "GiveItem" & x)
            Shop(y).TradeItem(x).GiveValue = GetVar(Filename, "SHOP" & y, "GiveValue" & x)
            Shop(y).TradeItem(x).GetItem = GetVar(Filename, "SHOP" & y, "GetItem" & x)
            Shop(y).TradeItem(x).GetValue = GetVar(Filename, "SHOP" & y, "GetValue" & x)
        Next x
    
        SetStatusProgress frmLoad.pBar.Value + 1
        DoEvents
    Next y
End Sub

Sub CheckShops()
    If Not FileExist("data\shops.ini") Then
        Call SaveShops
    End If
End Sub

Sub SaveSpell(ByVal SpellNum As Long)
Dim Filename As String
Dim i As Long

    Filename = App.Path & "\data\spells.ini"
    
    Call PutVar(Filename, "SPELL" & SpellNum, "Name", Trim(Spell(SpellNum).Name))
    Call PutVar(Filename, "SPELL" & SpellNum, "ClassReq", Trim(Spell(SpellNum).ClassReq))
    Call PutVar(Filename, "SPELL" & SpellNum, "Type", Trim(Spell(SpellNum).Type))
    Call PutVar(Filename, "SPELL" & SpellNum, "Data1", Trim(Spell(SpellNum).Data1))
    Call PutVar(Filename, "SPELL" & SpellNum, "Data2", Trim(Spell(SpellNum).Data2))
    Call PutVar(Filename, "SPELL" & SpellNum, "Data3", Trim(Spell(SpellNum).Data3))
End Sub

Sub SaveSpells()
Dim i As Long

    For i = 1 To MAX_SPELLS
        Call SaveSpell(i)
    Next i
End Sub

Sub LoadSpells()
Dim Filename As String
Dim i As Long

    Call CheckSpells
    
    Filename = App.Path & "\data\spells.ini"
    
    For i = 1 To MAX_SPELLS
        If Cancel_Load Then
            DestroyServer
            Exit Sub
        End If
        
        Spell(i).Name = GetVar(Filename, "SPELL" & i, "Name")
        Spell(i).ClassReq = Val(GetVar(Filename, "SPELL" & i, "ClassReq"))
        Spell(i).Type = Val(GetVar(Filename, "SPELL" & i, "Type"))
        Spell(i).Data1 = Val(GetVar(Filename, "SPELL" & i, "Data1"))
        Spell(i).Data2 = Val(GetVar(Filename, "SPELL" & i, "Data2"))
        Spell(i).Data3 = Val(GetVar(Filename, "SPELL" & i, "Data3"))
        
        SetStatusProgress frmLoad.pBar.Value + 1
        DoEvents
    Next i
End Sub

Sub CheckSpells()
    If Not FileExist("data\spells.ini") Then
        Call SaveSpells
    End If
End Sub

Sub SaveNpcs()
Dim i As Long
    
    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next i
End Sub

Sub SaveNpc(ByVal NpcNum As Long)
Dim Filename As String

    Filename = App.Path & "\data\npcs.ini"
    
    Call PutVar(Filename, "NPC" & NpcNum, "Name", Trim(Npc(NpcNum).Name))
    Call PutVar(Filename, "NPC" & NpcNum, "AttackSay", Trim(Npc(NpcNum).AttackSay))
    Call PutVar(Filename, "NPC" & NpcNum, "Sprite", Trim(Npc(NpcNum).Sprite))
    Call PutVar(Filename, "NPC" & NpcNum, "SpawnSecs", Trim(Npc(NpcNum).SpawnSecs))
    Call PutVar(Filename, "NPC" & NpcNum, "Behavior", Trim(Npc(NpcNum).Behavior))
    Call PutVar(Filename, "NPC" & NpcNum, "Range", Trim(Npc(NpcNum).Range))
    Call PutVar(Filename, "NPC" & NpcNum, "DropChance", Trim(Npc(NpcNum).DropChance))
    Call PutVar(Filename, "NPC" & NpcNum, "DropItem", Trim(Npc(NpcNum).DropItem))
    Call PutVar(Filename, "NPC" & NpcNum, "DropItemValue", Trim(Npc(NpcNum).DropItemValue))
    Call PutVar(Filename, "NPC" & NpcNum, "STR", Trim(Npc(NpcNum).STR))
    Call PutVar(Filename, "NPC" & NpcNum, "DEF", Trim(Npc(NpcNum).DEF))
    Call PutVar(Filename, "NPC" & NpcNum, "SPEED", Trim(Npc(NpcNum).SPEED))
    Call PutVar(Filename, "NPC" & NpcNum, "MAGI", Trim(Npc(NpcNum).MAGI))
End Sub

Sub LoadNpcs()
On Error Resume Next

Dim Filename As String
Dim i As Long

    Call CheckNpcs
    
    Filename = App.Path & "\data\npcs.ini"
    
    For i = 1 To MAX_NPCS
        If Cancel_Load Then
            DestroyServer
            Exit Sub
        End If
    
        Npc(i).Name = GetVar(Filename, "NPC" & i, "Name")
        Npc(i).AttackSay = GetVar(Filename, "NPC" & i, "AttackSay")
        Npc(i).Sprite = GetVar(Filename, "NPC" & i, "Sprite")
        Npc(i).SpawnSecs = GetVar(Filename, "NPC" & i, "SpawnSecs")
        Npc(i).Behavior = GetVar(Filename, "NPC" & i, "Behavior")
        Npc(i).Range = GetVar(Filename, "NPC" & i, "Range")
        Npc(i).DropChance = GetVar(Filename, "NPC" & i, "DropChance")
        Npc(i).DropItem = GetVar(Filename, "NPC" & i, "DropItem")
        Npc(i).DropItemValue = GetVar(Filename, "NPC" & i, "DropItemValue")
        Npc(i).STR = GetVar(Filename, "NPC" & i, "STR")
        Npc(i).DEF = GetVar(Filename, "NPC" & i, "DEF")
        Npc(i).SPEED = GetVar(Filename, "NPC" & i, "SPEED")
        Npc(i).MAGI = GetVar(Filename, "NPC" & i, "MAGI")
        
        SetStatusProgress frmLoad.pBar.Value + 1
        DoEvents
    Next i
End Sub

Sub CheckNpcs()
    If Not FileExist("data\npcs.ini") Then
        Call SaveNpcs
    End If
End Sub

Sub SaveMap(ByVal MapNum As Long)
Dim Filename As String
Dim f As Long

    Filename = App.Path & "\maps\map" & MapNum & ".dat"
        
    f = FreeFile
    Open Filename For Binary As #f
        Put #f, , Map(MapNum)
    Close #f
End Sub

Sub SaveMaps()
Dim Filename As String
Dim i As Long
Dim f As Long

    For i = 1 To MAX_MAPS
        Call SaveMap(i)
    Next i
End Sub

Sub LoadMaps()
Dim Filename As String
Dim i As Long
Dim f As Long

    Call CheckMaps
    
    For i = 1 To MAX_MAPS
        If Cancel_Load Then
            DestroyServer
            Exit Sub
        End If
        
        Filename = App.Path & "\maps\map" & i & ".dat"
        
        f = FreeFile
        Open Filename For Binary As #f
            Get #f, , Map(i)
        Close #f
        
        SetStatusProgress frmLoad.pBar.Value + 1
        DoEvents
    Next i
End Sub

Sub ConvertOldMapsToNew()
Dim Filename As String
Dim i As Long
Dim f As Long
Dim x As Long, y As Long
Dim OldMap As OldMapRec
Dim NewMap As MapRec

    For i = 1 To MAX_MAPS
        Filename = App.Path & "\maps\map" & i & ".dat"
        
        ' Get the old file
        f = FreeFile
        Open Filename For Binary As #f
            Get #f, , OldMap
        Close #f
        
        ' Delete the old file
        Call Kill(Filename)
        
        ' Convert
        NewMap.Name = OldMap.Name
        NewMap.Revision = OldMap.Revision + 1
        NewMap.Moral = OldMap.Moral
        NewMap.Up = OldMap.Up
        NewMap.Down = OldMap.Down
        NewMap.Left = OldMap.Left
        NewMap.Right = OldMap.Right
        NewMap.Music = OldMap.Music
        NewMap.BootMap = OldMap.BootMap
        NewMap.BootX = OldMap.BootX
        NewMap.BootY = OldMap.BootY
        NewMap.Shop = OldMap.Shop
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                NewMap.Tile(x, y).Ground = OldMap.Tile(x, y).Ground
                NewMap.Tile(x, y).Mask = OldMap.Tile(x, y).Mask
                NewMap.Tile(x, y).Anim = OldMap.Tile(x, y).Anim
                NewMap.Tile(x, y).Fringe = OldMap.Tile(x, y).Fringe
                NewMap.Tile(x, y).Type = OldMap.Tile(x, y).Type
                NewMap.Tile(x, y).Data1 = OldMap.Tile(x, y).Data1
                NewMap.Tile(x, y).Data2 = OldMap.Tile(x, y).Data2
                NewMap.Tile(x, y).Data3 = OldMap.Tile(x, y).Data3
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            NewMap.Npc(x) = OldMap.Npc(x)
        Next x
        
        ' Set new values to 0 or null
        NewMap.Indoors = NO
        
        ' Save the new map
        f = FreeFile
        Open Filename For Binary As #f
            Put #f, , NewMap
        Close #f
    Next i
End Sub

Sub CheckMaps()
    Dim Filename As String
    Dim i As Long

    For i = 1 To MAX_MAPS
        Filename = "maps\map" & i & ".dat"
        
        ' Check to see if map exists, if it doesn't, create it.
        If Not FileExist(Filename) Then
            Call SaveMap(i)
        End If
    Next i
End Sub

Sub AddLog(ByVal Text As String, ByVal FN As String)
Dim Filename As String
Dim f As Long

    If ServerLog = True Then
        Filename = App.Path & "\logs\" & FN
    
        If Not FileExist(FN) Then
            f = FreeFile
            Open Filename For Output As #f
            Close #f
        End If
    
        f = FreeFile
        Open Filename For Append As #f
            Print #f, Time & ": " & Text
        Close #f
    End If
End Sub

Sub BanIndex(ByVal BanPlayerIndex As Long, ByVal BannedByIndex As Long)
Dim Filename, IP As String
Dim f As Long, i As Long

    Filename = App.Path & "\data\banlist.txt"
    
    ' Make sure the file exists
    If Not FileExist("data\banlist.txt") Then
        f = FreeFile
        Open Filename For Output As #f
        Close #f
    End If
    
    ' Cut off last portion of ip
    IP = GetPlayerIP(BanPlayerIndex)
            
    For i = Len(IP) To 1 Step -1
        If Mid$(IP, i, 1) = "." Then
            Exit For
        End If
    Next i
    IP = Mid$(IP, 1, i)
            
    f = FreeFile
    Open Filename For Append As #f
        Print #f, IP & "," & GetPlayerName(BannedByIndex)
    Close #f
    
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & GAME_NAME & " by " & GetPlayerName(BannedByIndex) & "!", White)
    Call AddLog(GetPlayerName(BannedByIndex) & " has banned " & GetPlayerName(BanPlayerIndex) & ".", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned by " & GetPlayerName(BannedByIndex) & "!")
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
        If Trim(LCase(s)) <> Trim(LCase(Name)) Then
            Print #f2, s
        End If
    Loop
    
    Close #f1
    Close #f2
    
    Call Kill(App.Path & "\accounts\chartemp.txt")
End Sub
