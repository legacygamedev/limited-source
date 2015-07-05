Attribute VB_Name = "modDatabase"
Option Explicit

' Text API
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As String, ByVal lpfilename As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nsize As Long, ByVal lpfilename As String) As Long

'For Clear functions
Private Declare Sub ZeroMemory Lib "Kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

Public Sub HandleError(ByVal procName As String, ByVal contName As String, ByVal erNumber, ByVal erDesc, ByVal erSource, ByVal erHelpContext)
Dim filename As String
    filename = App.Path & "\data files\logs\errors.txt"
    Open filename For Append As #1
        Print #1, "The following error occured at '" & procName & "' in '" & contName & "'."
        Print #1, "Run-time error '" & erNumber & "': " & erDesc & "."
        Print #1, ""
    Close #1
End Sub

Public Sub ChkDir(ByVal tDir As String, ByVal tName As String)
    If LCase$(Dir(tDir & tName, vbDirectory)) <> tName Then Call MkDir(tDir & tName)
End Sub

' Outputs string to text file
Sub AddLog(ByVal Text As String, ByVal FN As String)
    Dim filename As String
    Dim F As Long

    If ServerLog Then
        filename = App.Path & "\data\logs\" & FN

        If Not FileExist(filename, True) Then
            F = FreeFile
            Open filename For Output As #F
            Close #F
        End If

        F = FreeFile
        Open filename For Append As #F
        Print #F, Time & ": " & Text
        Close #F
    End If

End Sub

' gets a string from a text file
Public Function GetVar(File As String, Header As String, Var As String) As String
    Dim sSpaces As String   ' Max string length
    Dim szReturn As String  ' Return default value if not found
    szReturn = vbNullString
    sSpaces = Space$(5000)
    Call GetPrivateProfileString$(Header, Var, szReturn, sSpaces, Len(sSpaces), File)
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

' writes a variable to a text file
Public Sub PutVar(File As String, Header As String, Var As String, Value As String)
    Call WritePrivateProfileString$(Header, Var, Value, File)
End Sub

Public Function FileExist(ByVal filename As String, Optional RAW As Boolean = False) As Boolean

    If Not RAW Then
        If LenB(Dir(App.Path & "\" & filename)) > 0 Then
            FileExist = True
        End If

    Else

        If LenB(Dir(filename)) > 0 Then
            FileExist = True
        End If
    End If

End Function

Public Sub SaveOptions()
    
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Game_Name", Options.Game_Name
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Port", STR(Options.Port)
    PutVar App.Path & "\data\options.ini", "OPTIONS", "MOTD", Options.MOTD
    PutVar App.Path & "\data\options.ini", "OPTIONS", "Website", Options.Website
    
End Sub

Public Sub LoadOptions()
    
    Options.Game_Name = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Game_Name")
    Options.Port = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Port")
    Options.MOTD = GetVar(App.Path & "\data\options.ini", "OPTIONS", "MOTD")
    Options.Website = GetVar(App.Path & "\data\options.ini", "OPTIONS", "Website")
    
End Sub

Public Sub ToggleMute(ByVal index As Long)
    ' exit out for rte9
    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub

    ' toggle the player's mute
    If Player(index).isMuted = 1 Then
        Player(index).isMuted = 0
        ' Let them know
        PlayerMsg index, "You have been unmuted and can now talk in global.", BrightGreen
        TextAdd GetPlayerName(index) & " has been unmuted."
    Else
        Player(index).isMuted = 1
        ' Let them know
        PlayerMsg index, "You have been muted and can no longer talk in global.", BrightRed
        TextAdd GetPlayerName(index) & " has been muted."
    End If
    
    ' save the player
    SavePlayer index
End Sub

Public Sub BanIndex(ByVal BanPlayerIndex As Long)
Dim filename As String, IP As String, F As Long, i As Long

    ' Add banned to the player's index
    Player(BanPlayerIndex).isBanned = 1
    SavePlayer BanPlayerIndex

    ' IP banning
    filename = App.Path & "\data\banlist_ip.txt"

    ' Make sure the file exists
    If Not FileExist(filename, True) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    ' Print the IP in the ip ban list
    IP = GetPlayerIP(BanPlayerIndex)
    F = FreeFile
    Open filename For Append As #F
        Print #F, IP
    Close #F
    
    ' Tell them they're banned
    Call GlobalMsg(GetPlayerName(BanPlayerIndex) & " has been banned from " & Options.Game_Name & ".", White)
    Call AddLog(GetPlayerName(BanPlayerIndex) & " has been banned.", ADMIN_LOG)
    Call AlertMsg(BanPlayerIndex, "You have been banned.")
End Sub

Public Function isBanned_IP(ByVal IP As String) As Boolean
Dim filename As String, fIP As String, F As Long
    
    filename = App.Path & "\data\banlist_ip.txt"

    ' Check if file exists
    If Not FileExist(filename, True) Then
        F = FreeFile
        Open filename For Output As #F
        Close #F
    End If

    F = FreeFile
    Open filename For Input As #F

    Do While Not EOF(F)
        Input #F, fIP

        ' Is banned?
        If Trim$(LCase$(fIP)) = Trim$(LCase$(Mid$(IP, 1, Len(fIP)))) Then
            isBanned_IP = True
            Close #F
            Exit Function
        End If
    Loop

    Close #F
End Function

Public Function isBanned_Account(ByVal index As Long) As Boolean
    If Player(index).isBanned = 1 Then
        isBanned_Account = True
    Else
        isBanned_Account = False
    End If
End Function

' **************
' ** Accounts **
' **************
Function AccountExist(ByVal Name As String) As Boolean
    Dim filename As String
    filename = App.Path & "\data\accounts\" & Trim(Name) & ".ini"

    If FileExist(filename, True) Then
        AccountExist = True
    End If

End Function

Function PasswordOK(ByVal Name As String, ByVal Password As String) As Boolean
Dim filename As String
Dim RightPassword As String

    If AccountExist(Name) Then
        filename = App.Path & "\data\accounts\" & Trim$(Name) & ".ini"
        
        RightPassword = GetVar(filename, "ACCOUNT", "Password")

        If UCase$(Trim$(Password)) = UCase$(Trim$(RightPassword)) Then
            PasswordOK = True
        End If
    End If

End Function

Sub AddAccount(ByVal index As Long, ByVal Name As String, ByVal Password As String)
    Dim i As Long
    
    ClearPlayer index
    
    Player(index).Login = Name
    Player(index).Password = Password

    Call SavePlayer(index)
End Sub

Sub DeleteName(ByVal Name As String)
    Dim f1 As Long
    Dim f2 As Long
    Dim s As String
    Call FileCopy(App.Path & "\data\accounts\_charlist.txt", App.Path & "\data\accounts\_chartemp.txt")
    ' Destroy name from charlist
    f1 = FreeFile
    Open App.Path & "\data\accounts\_chartemp.txt" For Input As #f1
    f2 = FreeFile
    Open App.Path & "\data\accounts\_charlist.txt" For Output As #f2

    Do While Not EOF(f1)
        Input #f1, s

        If Trim$(LCase$(s)) <> Trim$(LCase$(Name)) Then
            Print #f2, s
        End If

    Loop

    Close #f1
    Close #f2
    Call Kill(App.Path & "\data\accounts\chartemp.txt")
End Sub

' ****************
' ** Characters **
' ****************
Function CharExist(ByVal index As Long) As Boolean
    If LenB(Trim$(Player(index).Name)) > 0 Then
        CharExist = True
    End If
End Function

Sub AddChar(ByVal index As Long, ByVal Name As String, ByVal Sex As Byte, ByVal ClassNum As Long, ByVal Sprite As Long)
    Dim F As Long
    Dim n As Long
    Dim spritecheck As Boolean

    If LenB(Trim$(Player(index).Name)) = 0 Then
        
        spritecheck = False
        
        Player(index).Name = Name
        Player(index).Sex = Sex
        Player(index).Class = ClassNum
        
        If Player(index).Sex = SEX_MALE Then
            Player(index).Sprite = Class(ClassNum).MaleSprite(Sprite)
        Else
            Player(index).Sprite = Class(ClassNum).FemaleSprite(Sprite)
        End If

        Player(index).Level = 1

        For n = 1 To Stats.Stat_Count - 1
            Player(index).Stat(n) = Class(ClassNum).Stat(n)
        Next n

        Player(index).Dir = DIR_DOWN
        Player(index).Map = START_MAP
        Player(index).x = START_X
        Player(index).y = START_Y
        Player(index).Dir = DIR_DOWN
        Player(index).Vital(Vitals.HP) = GetPlayerMaxVital(index, Vitals.HP)
        Player(index).Vital(Vitals.MP) = GetPlayerMaxVital(index, Vitals.MP)
        
        ' set starter equipment
        If Class(ClassNum).startItemCount > 0 Then
            For n = 1 To Class(ClassNum).startItemCount
                If Class(ClassNum).StartItem(n) > 0 Then
                    ' item exist?
                    If Len(Trim$(Item(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Inv(n).Num = Class(ClassNum).StartItem(n)
                        Player(index).Inv(n).Value = Class(ClassNum).StartValue(n)
                    End If
                End If
            Next
        End If
        
        ' set start spells
        If Class(ClassNum).startSpellCount > 0 Then
            For n = 1 To Class(ClassNum).startSpellCount
                If Class(ClassNum).StartSpell(n) > 0 Then
                    ' spell exist?
                    If Len(Trim$(Spell(Class(ClassNum).StartItem(n)).Name)) > 0 Then
                        Player(index).Spell(n).Spell = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).Slot = Class(ClassNum).StartSpell(n)
                        Player(index).Hotbar(n).sType = 2 ' spells
                    End If
                End If
            Next
        End If
        
        ' Append name to file
        F = FreeFile
        Open App.Path & "\data\accounts\_charlist.txt" For Append As #F
        Print #F, Name
        Close #F
        Call SavePlayer(index)
        Exit Sub
    End If

End Sub

Function FindChar(ByVal Name As String) As Boolean
    Dim F As Long
    Dim s As String
    F = FreeFile
    Open App.Path & "\data\accounts\_charlist.txt" For Input As #F

    Do While Not EOF(F)
        Input #F, s

        If Trim$(LCase$(s)) = Trim$(LCase$(Name)) Then
            FindChar = True
            Close #F
            Exit Function
        End If

    Loop

    Close #F
End Function

' *************
' ** Players **
' *************
Sub SaveAllPlayersOnline()
    Dim i As Long

    For i = 1 To Player_HighIndex

        If IsPlaying(i) Then
            Call SavePlayer(i)
            Call SaveBank(i)
        End If

    Next

End Sub

Sub SavePlayer(ByVal index As Long)
Dim filename As String, i As Long

    If index <= 0 Or index > MAX_PLAYERS Then Exit Sub
    
    ' the file
    filename = App.Path & "\data\accounts\" & Trim$(Player(index).Login) & ".ini"
    
    ' General
    PutVar filename, "ACCOUNT", "Login", Trim$(Player(index).Login)
    PutVar filename, "ACCOUNT", "Password", Trim$(Player(index).Password)
    PutVar filename, "ACCOUNT", "Name", Trim$(Player(index).Name)
    PutVar filename, "ACCOUNT", "Sex", Val(Player(index).Sex)
    PutVar filename, "ACCOUNT", "Class", Val(Player(index).Class)
    PutVar filename, "ACCOUNT", "Sprite", Val(Player(index).Sprite)
    PutVar filename, "ACCOUNT", "Level", Val(Player(index).Level)
    PutVar filename, "ACCOUNT", "exp", Val(Player(index).exp)
    PutVar filename, "ACCOUNT", "Access", Val(Player(index).Access)
    PutVar filename, "ACCOUNT", "PK", Val(Player(index).PK)
    
    ' Vitals
    For i = 1 To Vitals.Vital_Count - 1
        PutVar filename, "ACCOUNT", "Vital" & i, Val(Player(index).Vital(i))
    Next
    
    ' Stats
    For i = 1 To Stats.Stat_Count - 1
        PutVar filename, "ACCOUNT", "Stat" & i, Val(Player(index).Stat(i))
    Next
    PutVar filename, "ACCOUNT", "Points", Val(Player(index).POINTS)

    ' Equipment
    For i = 1 To Equipment.Equipment_Count - 1
        PutVar filename, "ACCOUNT", "Equipment" & i, Val(Player(index).Equipment(i))
    Next
    
    ' Inventory
    For i = 1 To MAX_INV
        PutVar filename, "ACCOUNT", "InvNum" & i, Val(Player(index).Inv(i).Num)
        PutVar filename, "ACCOUNT", "InvValue" & i, Val(Player(index).Inv(i).Value)
        PutVar filename, "ACCOUNT", "InvBound" & i, Val(Player(index).Inv(i).Bound)
    Next
    
    ' Spells
    For i = 1 To MAX_PLAYER_SPELLS
        PutVar filename, "ACCOUNT", "Spell" & i, Val(Player(index).Spell(i).Spell)
        PutVar filename, "ACCOUNT", "SpellUses" & i, Val(Player(index).Spell(i).Uses)
    Next
    
    ' Hotbar
    For i = 1 To MAX_HOTBAR
        PutVar filename, "ACCOUNT", "HotbarSlot" & i, Val(Player(index).Hotbar(i).Slot)
        PutVar filename, "ACCOUNT", "HotbarType" & i, Val(Player(index).Hotbar(i).sType)
    Next
    
    ' Position
    PutVar filename, "ACCOUNT", "Map", Val(Player(index).Map)
    PutVar filename, "ACCOUNT", "X", Val(Player(index).x)
    PutVar filename, "ACCOUNT", "Y", Val(Player(index).y)
    PutVar filename, "ACCOUNT", "Dir", Val(Player(index).Dir)
    
    ' Tutorial
    PutVar filename, "ACCOUNT", "TutorialState", Val(Player(index).TutorialState)
    
    ' Banned
    PutVar filename, "ACCOUNT", "isBanned", Val(Player(index).isBanned)
    PutVar filename, "ACCOUNT", "isMuted", Val(Player(index).isMuted)
End Sub

Sub LoadPlayer(ByVal index As Long, ByVal Name As String)
Dim filename As String, i As Long

    If Trim$(Name) = vbNullString Then Exit Sub

    ' clear player
    Call ClearPlayer(index)
    
    ' the file
    filename = App.Path & "\data\accounts\" & Trim$(Name) & ".ini"

    ' General
    Player(index).Login = GetVar(filename, "ACCOUNT", "Login")
    Player(index).Password = GetVar(filename, "ACCOUNT", "Password")
    Player(index).Name = GetVar(filename, "ACCOUNT", "Name")
    Player(index).Sex = Val(GetVar(filename, "ACCOUNT", "Sex"))
    Player(index).Class = Val(GetVar(filename, "ACCOUNT", "Class"))
    Player(index).Sprite = Val(GetVar(filename, "ACCOUNT", "Sprite"))
    Player(index).Level = Val(GetVar(filename, "ACCOUNT", "Level"))
    Player(index).exp = Val(GetVar(filename, "ACCOUNT", "Exp"))
    Player(index).Access = Val(GetVar(filename, "ACCOUNT", "Access"))
    Player(index).PK = Val(GetVar(filename, "ACCOUNT", "PK"))
    
    ' Vitals
    For i = 1 To Vitals.Vital_Count - 1
        Player(index).Vital(i) = Val(GetVar(filename, "ACCOUNT", "Vital" & i))
    Next
    
    ' Stats
    For i = 1 To Stats.Stat_Count - 1
        Player(index).Stat(i) = Val(GetVar(filename, "ACCOUNT", "Stat" & i))
    Next
    Player(index).POINTS = Val(GetVar(filename, "ACCOUNT", "Points"))

    ' Equipment
    For i = 1 To Equipment.Equipment_Count - 1
        Player(index).Equipment(i) = Val(GetVar(filename, "ACCOUNT", "Equipment" & i))
    Next
    
    ' Inventory
    For i = 1 To MAX_INV
        Player(index).Inv(i).Num = Val(GetVar(filename, "ACCOUNT", "InvNum" & i))
        Player(index).Inv(i).Value = Val(GetVar(filename, "ACCOUNT", "InvValue" & i))
        Player(index).Inv(i).Bound = Val(GetVar(filename, "ACCOUNT", "InvBound" & i))
    Next
    
    ' Spells
    For i = 1 To MAX_PLAYER_SPELLS
        Player(index).Spell(i).Spell = Val(GetVar(filename, "ACCOUNT", "Spell" & i))
        Player(index).Spell(i).Uses = Val(GetVar(filename, "ACCOUNT", "SpellUses" & i))
    Next
    
    ' Hotbar
    For i = 1 To MAX_HOTBAR
        Player(index).Hotbar(i).Slot = Val(GetVar(filename, "ACCOUNT", "HotbarSlot" & i))
        Player(index).Hotbar(i).sType = Val(GetVar(filename, "ACCOUNT", "HotbarType" & i))
    Next
    
    ' Position
    Player(index).Map = Val(GetVar(filename, "ACCOUNT", "Map"))
    Player(index).x = Val(GetVar(filename, "ACCOUNT", "X"))
    Player(index).y = Val(GetVar(filename, "ACCOUNT", "Y"))
    Player(index).Dir = Val(GetVar(filename, "ACCOUNT", "Dir"))
    
    ' Tutorial
    Player(index).TutorialState = Val(GetVar(filename, "ACCOUNT", "TutorialState"))
    
    ' Banned
    Player(index).isBanned = Val(GetVar(filename, "ACCOUNT", "isBanned"))
    Player(index).isMuted = Val(GetVar(filename, "ACCOUNT", "isMuted"))
End Sub

Sub ClearPlayer(ByVal index As Long)
    Dim i As Long
    
    Call ZeroMemory(ByVal VarPtr(TempPlayer(index)), LenB(TempPlayer(index)))
    Set TempPlayer(index).Buffer = New clsBuffer
    
    Call ZeroMemory(ByVal VarPtr(Player(index)), LenB(Player(index)))
    Player(index).Login = vbNullString
    Player(index).Password = vbNullString
    Player(index).Name = vbNullString
    Player(index).Class = 1

    frmServer.lvwInfo.ListItems(index).SubItems(1) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(2) = vbNullString
    frmServer.lvwInfo.ListItems(index).SubItems(3) = vbNullString
End Sub

' *************
' ** Classes **
' *************
Public Sub CreateClassesINI()
    Dim filename As String
    Dim File As String
    filename = App.Path & "\data\classes.ini"
    Max_Classes = 2

    If Not FileExist(filename, True) Then
        File = FreeFile
        Open filename For Output As File
        Print #File, "[INIT]"
        Print #File, "MaxClasses=" & Max_Classes
        Close File
    End If

End Sub

Sub LoadClasses()
    Dim filename As String
    Dim i As Long, n As Long
    Dim tmpSprite As String
    Dim tmpArray() As String
    Dim startItemCount As Long, startSpellCount As Long
    Dim x As Long

    If CheckClasses Then
        ReDim Class(1 To Max_Classes)
        Call SaveClasses
    Else
        filename = App.Path & "\data\classes.ini"
        Max_Classes = Val(GetVar(filename, "INIT", "MaxClasses"))
        ReDim Class(1 To Max_Classes)
    End If

    Call ClearClasses

    For i = 1 To Max_Classes
        Class(i).Name = GetVar(filename, "CLASS" & i, "Name")
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "MaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).MaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).MaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' read string of sprites
        tmpSprite = GetVar(filename, "CLASS" & i, "FemaleSprite")
        ' split into an array of strings
        tmpArray() = Split(tmpSprite, ",")
        ' redim the class sprite array
        ReDim Class(i).FemaleSprite(0 To UBound(tmpArray))
        ' loop through converting strings to values and store in the sprite array
        For n = 0 To UBound(tmpArray)
            Class(i).FemaleSprite(n) = Val(tmpArray(n))
        Next
        
        ' continue
        Class(i).Stat(Stats.Strength) = Val(GetVar(filename, "CLASS" & i, "Strength"))
        Class(i).Stat(Stats.Endurance) = Val(GetVar(filename, "CLASS" & i, "Endurance"))
        Class(i).Stat(Stats.Intelligence) = Val(GetVar(filename, "CLASS" & i, "Intelligence"))
        Class(i).Stat(Stats.Agility) = Val(GetVar(filename, "CLASS" & i, "Agility"))
        Class(i).Stat(Stats.Willpower) = Val(GetVar(filename, "CLASS" & i, "Willpower"))
        
        ' how many starting items?
        startItemCount = Val(GetVar(filename, "CLASS" & i, "StartItemCount"))
        If startItemCount > 0 Then ReDim Class(i).StartItem(1 To startItemCount)
        If startItemCount > 0 Then ReDim Class(i).StartValue(1 To startItemCount)
        
        ' loop for items & values
        Class(i).startItemCount = startItemCount
        If startItemCount >= 1 And startItemCount <= MAX_INV Then
            For x = 1 To startItemCount
                Class(i).StartItem(x) = Val(GetVar(filename, "CLASS" & i, "StartItem" & x))
                Class(i).StartValue(x) = Val(GetVar(filename, "CLASS" & i, "StartValue" & x))
            Next
        End If
        
        ' how many starting spells?
        startSpellCount = Val(GetVar(filename, "CLASS" & i, "StartSpellCount"))
        If startSpellCount > 0 Then ReDim Class(i).StartSpell(1 To startSpellCount)
        
        ' loop for spells
        Class(i).startSpellCount = startSpellCount
        If startSpellCount >= 1 And startSpellCount <= MAX_INV Then
            For x = 1 To startSpellCount
                Class(i).StartSpell(x) = Val(GetVar(filename, "CLASS" & i, "StartSpell" & x))
            Next
        End If
    Next

End Sub

Sub SaveClasses()
    Dim filename As String
    Dim i As Long
    Dim x As Long
    
    filename = App.Path & "\data\classes.ini"

    For i = 1 To Max_Classes
        Call PutVar(filename, "CLASS" & i, "Name", Trim$(Class(i).Name))
        Call PutVar(filename, "CLASS" & i, "Maleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Femaleprite", "1")
        Call PutVar(filename, "CLASS" & i, "Strength", STR(Class(i).Stat(Stats.Strength)))
        Call PutVar(filename, "CLASS" & i, "Endurance", STR(Class(i).Stat(Stats.Endurance)))
        Call PutVar(filename, "CLASS" & i, "Intelligence", STR(Class(i).Stat(Stats.Intelligence)))
        Call PutVar(filename, "CLASS" & i, "Agility", STR(Class(i).Stat(Stats.Agility)))
        Call PutVar(filename, "CLASS" & i, "Willpower", STR(Class(i).Stat(Stats.Willpower)))
        ' loop for items & values
        For x = 1 To UBound(Class(i).StartItem)
            Call PutVar(filename, "CLASS" & i, "StartItem" & x, STR(Class(i).StartItem(x)))
            Call PutVar(filename, "CLASS" & i, "StartValue" & x, STR(Class(i).StartValue(x)))
        Next
        ' loop for spells
        For x = 1 To UBound(Class(i).StartSpell)
            Call PutVar(filename, "CLASS" & i, "StartSpell" & x, STR(Class(i).StartSpell(x)))
        Next
    Next

End Sub

Function CheckClasses() As Boolean
    Dim filename As String
    filename = App.Path & "\data\classes.ini"

    If Not FileExist(filename, True) Then
        Call CreateClassesINI
        CheckClasses = True
    End If

End Function

Sub ClearClasses()
    Dim i As Long

    For i = 1 To Max_Classes
        Call ZeroMemory(ByVal VarPtr(Class(i)), LenB(Class(i)))
        Class(i).Name = vbNullString
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

Sub SaveItem(ByVal itemNum As Long)
    Dim filename As String
    Dim F  As Long
    filename = App.Path & "\data\items\item" & itemNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Item(itemNum)
    Close #F
End Sub

Sub LoadItems()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckItems

    For i = 1 To MAX_ITEMS
        filename = App.Path & "\data\Items\Item" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Item(i)
        Close #F
    Next

End Sub

Sub CheckItems()
    Dim i As Long

    For i = 1 To MAX_ITEMS

        If Not FileExist("\Data\Items\Item" & i & ".dat") Then
            Call SaveItem(i)
        End If

    Next

End Sub

Sub ClearItem(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Item(index)), LenB(Item(index)))
    Item(index).Name = vbNullString
    Item(index).Desc = vbNullString
    Item(index).Sound = "None."
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

Sub SaveShop(ByVal shopNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\shops\shop" & shopNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Shop(shopNum)
    Close #F
End Sub

Sub LoadShops()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckShops

    For i = 1 To MAX_SHOPS
        filename = App.Path & "\data\shops\shop" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Shop(i)
        Close #F
    Next

End Sub

Sub CheckShops()
    Dim i As Long

    For i = 1 To MAX_SHOPS

        If Not FileExist("\Data\shops\shop" & i & ".dat") Then
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
Sub SaveSpell(ByVal spellNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\spells\spells" & spellNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Spell(spellNum)
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
        filename = App.Path & "\data\spells\spells" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Spell(i)
        Close #F
    Next

End Sub

Sub CheckSpells()
    Dim i As Long

    For i = 1 To MAX_SPELLS

        If Not FileExist("\Data\spells\spells" & i & ".dat") Then
            Call SaveSpell(i)
        End If

    Next

End Sub

Sub ClearSpell(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Spell(index)), LenB(Spell(index)))
    Spell(index).Name = vbNullString
    Spell(index).LevelReq = 1 'Needs to be 1 for the spell editor
    Spell(index).Desc = vbNullString
    Spell(index).Sound = "None."
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
Sub SaveNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call SaveNpc(i)
    Next

End Sub

Sub SaveNpc(ByVal NpcNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\npcs\npc" & NpcNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Npc(NpcNum)
    Close #F
End Sub

Sub LoadNpcs()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Call CheckNpcs

    For i = 1 To MAX_NPCS
        filename = App.Path & "\data\npcs\npc" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Npc(i)
        Close #F
    Next

End Sub

Sub CheckNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS

        If Not FileExist("\Data\npcs\npc" & i & ".dat") Then
            Call SaveNpc(i)
        End If

    Next

End Sub

Sub ClearNpc(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Npc(index)), LenB(Npc(index)))
    Npc(index).Name = vbNullString
    Npc(index).AttackSay = vbNullString
    Npc(index).Sound = "None."
End Sub

Sub ClearNpcs()
    Dim i As Long

    For i = 1 To MAX_NPCS
        Call ClearNpc(i)
    Next

End Sub

' **********
' ** Resources **
' **********
Sub SaveResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call SaveResource(i)
    Next

End Sub

Sub SaveResource(ByVal ResourceNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\resources\resource" & ResourceNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Resource(ResourceNum)
    Close #F
End Sub

Sub LoadResources()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckResources

    For i = 1 To MAX_RESOURCES
        filename = App.Path & "\data\resources\resource" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Resource(i)
        Close #F
    Next

End Sub

Sub CheckResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Not FileExist("\Data\Resources\Resource" & i & ".dat") Then
            Call SaveResource(i)
        End If
    Next

End Sub

Sub ClearResource(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Resource(index)), LenB(Resource(index)))
    Resource(index).Name = vbNullString
    Resource(index).SuccessMessage = vbNullString
    Resource(index).EmptyMessage = vbNullString
    Resource(index).Sound = "None."
End Sub

Sub ClearResources()
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        Call ClearResource(i)
    Next
End Sub

' **********
' ** animations **
' **********
Sub SaveAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        Call SaveAnimation(i)
    Next

End Sub

Sub SaveAnimation(ByVal AnimationNum As Long)
    Dim filename As String
    Dim F As Long
    filename = App.Path & "\data\animations\animation" & AnimationNum & ".dat"
    F = FreeFile
    Open filename For Binary As #F
        Put #F, , Animation(AnimationNum)
    Close #F
End Sub

Sub LoadAnimations()
    Dim filename As String
    Dim i As Long
    Dim F As Long
    Dim sLen As Long
    
    Call CheckAnimations

    For i = 1 To MAX_ANIMATIONS
        filename = App.Path & "\data\animations\animation" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            Get #F, , Animation(i)
        Close #F
    Next

End Sub

Sub CheckAnimations()
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS

        If Not FileExist("\Data\animations\animation" & i & ".dat") Then
            Call SaveAnimation(i)
        End If

    Next

End Sub

Sub ClearAnimation(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Animation(index)), LenB(Animation(index)))
    Animation(index).Name = vbNullString
    Animation(index).Sound = "None."
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
Sub SaveMap(ByVal mapnum As Long)
    Dim filename As String
    Dim F As Long
    Dim x As Long
    Dim y As Long
    filename = App.Path & "\data\maps\map" & mapnum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
    Put #F, , Map(mapnum).Name
    Put #F, , Map(mapnum).Music
    Put #F, , Map(mapnum).Revision
    Put #F, , Map(mapnum).Moral
    Put #F, , Map(mapnum).Up
    Put #F, , Map(mapnum).Down
    Put #F, , Map(mapnum).Left
    Put #F, , Map(mapnum).Right
    Put #F, , Map(mapnum).BootMap
    Put #F, , Map(mapnum).BootX
    Put #F, , Map(mapnum).BootY
    Put #F, , Map(mapnum).MaxX
    Put #F, , Map(mapnum).MaxY

    For x = 0 To Map(mapnum).MaxX
        For y = 0 To Map(mapnum).MaxY
            Put #F, , Map(mapnum).Tile(x, y)
        Next
    Next

    For x = 1 To MAX_MAP_NPCS
        Put #F, , Map(mapnum).Npc(x)
    Next
    
    Put #F, , Map(mapnum).BossNpc
    Close #F
    
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
    Dim y As Long
    Call CheckMaps

    For i = 1 To MAX_MAPS
        filename = App.Path & "\data\maps\map" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
        Get #F, , Map(i).Name
        Get #F, , Map(i).Music
        Get #F, , Map(i).Revision
        Get #F, , Map(i).Moral
        Get #F, , Map(i).Up
        Get #F, , Map(i).Down
        Get #F, , Map(i).Left
        Get #F, , Map(i).Right
        Get #F, , Map(i).BootMap
        Get #F, , Map(i).BootX
        Get #F, , Map(i).BootY
        Get #F, , Map(i).MaxX
        Get #F, , Map(i).MaxY
        ' have to set the tile()
        ReDim Map(i).Tile(0 To Map(i).MaxX, 0 To Map(i).MaxY)

        For x = 0 To Map(i).MaxX
            For y = 0 To Map(i).MaxY
                Get #F, , Map(i).Tile(x, y)
            Next
        Next

        For x = 1 To MAX_MAP_NPCS
            Get #F, , Map(i).Npc(x)
            MapNpc(i).Npc(x).Num = Map(i).Npc(x)
        Next
        
        Get #F, , Map(i).BossNpc
        Close #F
        
        ClearTempTile i
        CacheResources i
        DoEvents
    Next
End Sub

Sub CheckMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS

        If Not FileExist("\Data\maps\map" & i & ".dat") Then
            Call SaveMap(i)
        End If

    Next

End Sub

Sub ClearMapItem(ByVal index As Long, ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(MapItem(mapnum, index)), LenB(MapItem(mapnum, index)))
    MapItem(mapnum, index).playerName = vbNullString
End Sub

Sub ClearMapItems()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_ITEMS
            Call ClearMapItem(x, y)
        Next
    Next

End Sub

Sub ClearMapNpc(ByVal index As Long, ByVal mapnum As Long)
    ReDim MapNpc(mapnum).Npc(1 To MAX_MAP_NPCS)
    Call ZeroMemory(ByVal VarPtr(MapNpc(mapnum).Npc(index)), LenB(MapNpc(mapnum).Npc(index)))
End Sub

Sub ClearMapNpcs()
    Dim x As Long
    Dim y As Long

    For y = 1 To MAX_MAPS
        For x = 1 To MAX_MAP_NPCS
            Call ClearMapNpc(x, y)
        Next
    Next

End Sub

Sub ClearMap(ByVal mapnum As Long)
    Call ZeroMemory(ByVal VarPtr(Map(mapnum)), LenB(Map(mapnum)))
    Map(mapnum).Name = vbNullString
    Map(mapnum).MaxX = MAX_MAPX
    Map(mapnum).MaxY = MAX_MAPY
    ReDim Map(mapnum).Tile(0 To Map(mapnum).MaxX, 0 To Map(mapnum).MaxY)
    ' Reset the values for if a player is on the map or not
    PlayersOnMap(mapnum) = NO
    ' Reset the map cache array for this map.
    MapCache(mapnum).Data = vbNullString
End Sub

Sub ClearMaps()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call ClearMap(i)
    Next

End Sub

Function GetClassName(ByVal ClassNum As Long) As String
    GetClassName = Trim$(Class(ClassNum).Name)
End Function

Function GetClassMaxVital(ByVal ClassNum As Long, ByVal Vital As Vitals) As Long
    Select Case Vital
        Case HP
            With Class(ClassNum)
                GetClassMaxVital = 100 + (.Stat(Endurance) * 5) + 2
            End With
        Case MP
            With Class(ClassNum)
                GetClassMaxVital = 30 + (.Stat(Intelligence) * 10) + 2
            End With
    End Select
End Function

Function GetClassStat(ByVal ClassNum As Long, ByVal Stat As Stats) As Long
    GetClassStat = Class(ClassNum).Stat(Stat)
End Function

Sub SaveBank(ByVal index As Long)
    Dim filename As String
    Dim F As Long
    
    filename = App.Path & "\data\banks\" & Trim$(Player(index).Login) & ".bin"
    
    F = FreeFile
    Open filename For Binary As #F
    Put #F, , Bank(index)
    Close #F
End Sub

Public Sub LoadBank(ByVal index As Long, ByVal Name As String)
    Dim filename As String
    Dim F As Long

    Call ClearBank(index)

    filename = App.Path & "\data\banks\" & Trim$(Name) & ".bin"
    
    If Not FileExist(filename, True) Then
        Call SaveBank(index)
        Exit Sub
    End If

    F = FreeFile
    Open filename For Binary As #F
        Get #F, , Bank(index)
    Close #F

End Sub

Sub ClearBank(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Bank(index)), LenB(Bank(index)))
End Sub

Sub ClearParty(ByVal partynum As Long)
    Call ZeroMemory(ByVal VarPtr(Party(partynum)), LenB(Party(partynum)))
End Sub

' ***********
' ** Convs **
' ***********
Sub SaveConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        Call SaveConv(i)
    Next
End Sub

Sub SaveConv(ByVal convNum As Long)
Dim filename As String
Dim i As Long, x As Long, F As Long
    
    filename = App.Path & "\data\convs\conv" & convNum & ".dat"
    F = FreeFile
    
    Open filename For Binary As #F
        With Conv(convNum)
            Put #F, , .Name
            Put #F, , .chatCount
            For i = 1 To .chatCount
                Put #F, , CLng(Len(.Conv(i).Conv))
                Put #F, , .Conv(i).Conv
                For x = 1 To 4
                    Put #F, , CLng(Len(.Conv(i).rText(x)))
                    Put #F, , .Conv(i).rText(x)
                    Put #F, , .Conv(i).rTarget(x)
                Next
                Put #F, , .Conv(i).Event
                Put #F, , .Conv(i).data1
                Put #F, , .Conv(i).data2
                Put #F, , .Conv(i).data3
            Next
        End With
    Close #F
End Sub

Sub LoadConvs()
Dim filename As String
Dim i As Long, n As Long, x As Long, F As Long
Dim sLen As Long
    
    Call CheckConvs

    For i = 1 To MAX_CONVS
        filename = App.Path & "\data\convs\conv" & i & ".dat"
        F = FreeFile
        Open filename For Binary As #F
            With Conv(i)
                Get #F, , .Name
                Get #F, , .chatCount
                If .chatCount > 0 Then ReDim .Conv(1 To .chatCount)
                For n = 1 To .chatCount
                    Get #F, , sLen
                    .Conv(n).Conv = Space$(sLen)
                    Get #F, , .Conv(n).Conv
                    For x = 1 To 4
                        Get #F, , sLen
                        .Conv(n).rText(x) = Space$(sLen)
                        Get #F, , .Conv(n).rText(x)
                        Get #F, , .Conv(n).rTarget(x)
                    Next
                    Get #F, , .Conv(n).Event
                    Get #F, , .Conv(n).data1
                    Get #F, , .Conv(n).data2
                    Get #F, , .Conv(n).data3
                Next
            End With
        Close #F
    Next
End Sub

Sub CheckConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        If Not FileExist("\data\convs\conv" & i & ".dat") Then
            Call SaveConv(i)
        End If
    Next
End Sub

Sub ClearConv(ByVal index As Long)
    Call ZeroMemory(ByVal VarPtr(Conv(index)), LenB(Conv(index)))
    Conv(index).Name = vbNullString
    ReDim Conv(index).Conv(1)
End Sub

Sub ClearConvs()
Dim i As Long

    For i = 1 To MAX_CONVS
        Call ClearConv(i)
    Next

End Sub
