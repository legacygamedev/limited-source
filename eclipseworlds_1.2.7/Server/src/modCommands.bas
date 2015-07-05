Attribute VB_Name = "modCommands"
Option Explicit

Function GetPlayerLogin(ByVal index As Long) As String

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerLogin = Trim$(Account(index).Login)
End Function

Sub SetPlayerLogin(ByVal index As Long, ByVal Login As String)
    Account(index).Login = Login
End Sub

Function GetPlayerPassword(ByVal index As Long) As String

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerPassword = Trim$(Account(index).Password)
End Function

Sub SetPlayerPassword(ByVal index As Long, ByVal Password As String)
    Account(index).Password = Password
End Sub

Function GetPlayerName(ByVal index As Long) As String

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Account(index).Chars(GetPlayerChar(index)).Name)
End Function

Sub SetPlayerName(ByVal index As Long, ByVal Name As String)
    Account(index).Chars(GetPlayerChar(index)).Name = Name
End Sub

Function GetPlayerChar(ByVal index As Byte) As Byte

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerChar = Account(index).CurrentChar
    
    If GetPlayerChar = 0 Then GetPlayerChar = 1
End Function

Sub SetPlayerChar(ByVal index As Long, ByVal Char As Byte)
    Account(index).CurrentChar = index
End Sub

Function GetPlayerGuildName(ByVal index As Long) As String
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerGuildName = Trim$(Guild(Account(index).Chars(GetPlayerChar(index)).Guild.index).Name)
End Function

Function GetPlayerGuild(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerGuild = Account(index).Chars(GetPlayerChar(index)).Guild.index
End Function

Sub SetPlayerGuild(ByVal index As Long, ByVal GuildNum As Long)
    Account(index).Chars(GetPlayerChar(index)).Guild.index = GuildNum
End Sub

Function GetPlayerGuildAccess(ByVal index As Long) As Byte

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerGuildAccess = Account(index).Chars(GetPlayerChar(index)).Guild.Access
End Function

Sub SetPlayerGuildAccess(ByVal index As Long, ByVal Access As Byte)
    Account(index).Chars(GetPlayerChar(index)).Guild.Access = Access
End Sub

Function GetPlayerClass(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Account(index).Chars(GetPlayerChar(index)).Class
End Function

Sub SetPlayerClass(ByVal index As Long, ByVal ClassNum As Long)
    Account(index).Chars(GetPlayerChar(index)).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Account(index).Chars(GetPlayerChar(index)).Sprite
End Function

Sub SetPlayerSprite(ByVal index As Long, ByVal Sprite As Long)
    Account(index).Chars(GetPlayerChar(index)).Sprite = Sprite
End Sub

Function GetPlayerTitle(ByVal index As Long, ByVal TitleNum As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).Title(MAX_TITLES)
    GetPlayerTitle = Account(index).Chars(GetPlayerChar(index)).Title(TitleNum)
End Function

Sub SetPlayerTitle(ByVal index As Long, ByVal Title As Long, ByVal TitleNum As Long)
    ReDim Preserve Account(index).Chars(GetPlayerChar(index)).Title(MAX_TITLES)
    Account(index).Chars(GetPlayerChar(index)).Title(Title) = TitleNum
End Sub

Function GetPlayerFace(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerFace = Account(index).Chars(GetPlayerChar(index)).Face
End Function

Sub SetPlayerFace(ByVal index As Long, ByVal Face As Long)
    Account(index).Chars(GetPlayerChar(index)).Face = Face
End Sub

Function GetPlayerLevel(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Account(index).Chars(GetPlayerChar(index)).Level
End Function

Function GetPlayerSkill(ByVal index As Long, ByVal SkillNum As Byte) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerSkill = Account(index).Chars(GetPlayerChar(index)).Skills(SkillNum).Level
End Function

Sub SetPlayerLevel(ByVal index As Long, ByVal Level As Byte, Optional ByVal PlusVal As Boolean = False)

    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    If Not PlusVal Then
        Account(index).Chars(GetPlayerChar(index)).Level = Level
    Else
        If Not Account(index).Chars(GetPlayerChar(index)).Level + Level > Options.MaxLevel And Not Account(index).Chars(GetPlayerChar(index)).Level + Level < 1 Then
            Account(index).Chars(GetPlayerChar(index)).Level = Account(index).Chars(GetPlayerChar(index)).Level + Level
        End If
    End If
End Sub

Sub SetPlayerSkill(ByVal index As Long, ByVal Level As Byte, ByVal SkillNum As Byte, Optional ByVal PlusVal As Boolean = False)
Dim i As Long, NPCNum As Long, Parse() As String

        If index < 1 Or index > MAX_PLAYERS Then Exit Sub
        
        For i = 1 To MAX_QUESTS
            Parse() = Split(HasQuestSkill(index, i, True), "|")
            If UBound(Parse()) > 0 Then
                NPCNum = Parse(0)
                If NPCNum > 0 Then
                    Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
                End If
            Else
                NPCNum = HasQuestSkill(index, i)
                If NPCNum > 0 Then
                        Call SendShowTaskCompleteOnNPC(index, NPCNum, True)
                End If
            End If
        Next i
        
    If Not PlusVal Then
        Account(index).Chars(GetPlayerChar(index)).Skills(SkillNum).Level = Level
    Else
        Account(index).Chars(GetPlayerChar(index)).Skills(SkillNum).Level = Account(index).Chars(GetPlayerChar(index)).Skills(SkillNum).Level + Level
    End If
End Sub

Function GetPlayerNextLevel(ByVal index As Long) As Long
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(index) + 1) ^ 3 - (6 * (GetPlayerLevel(index) + 1) ^ 2) + 17 * (GetPlayerLevel(index) + 1) - 12)
End Function

Function GetPlayerNextSkillLevel(ByVal index As Long, ByVal SkillNum As Byte) As Long
    GetPlayerNextSkillLevel = (50 / 3) * ((GetPlayerSkill(index, SkillNum) + 1) ^ 3 - (6 * (GetPlayerSkill(index, SkillNum) + 1) ^ 2) + 17 * (GetPlayerSkill(index, SkillNum) + 1) - 12)
End Function

Function GetPlayerExp(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Account(index).Chars(GetPlayerChar(index)).Exp
End Function

Function GetPlayerSkillExp(ByVal index As Long, ByVal SkillNum As Byte) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerSkillExp = Account(index).Chars(GetPlayerChar(index)).Skills(SkillNum).Exp
End Function

Sub SetPlayerExp(ByVal index As Long, ByVal Exp As Long, Optional ByVal PlusVal As Boolean = False)
    If Not PlusVal Then
        Account(index).Chars(GetPlayerChar(index)).Exp = Exp
    Else
        Account(index).Chars(GetPlayerChar(index)).Exp = Account(index).Chars(GetPlayerChar(index)).Exp + Exp
    End If
End Sub

Sub SetPlayerSkillExp(ByVal index As Long, ByVal Exp As Long, ByVal SkillNum As Byte, Optional ByVal PlusVal As Boolean = False)
    If Not PlusVal Then
        Account(index).Chars(GetPlayerChar(index)).Skills(SkillNum).Exp = Exp
    Else
        Account(index).Chars(GetPlayerChar(index)).Skills(SkillNum).Exp = Account(index).Chars(GetPlayerChar(index)).Skills(SkillNum).Exp + Exp
    End If
End Sub

Function GetPlayerAccess(ByVal index As Long) As Byte

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Account(index).Chars(GetPlayerChar(index)).Access
End Function

Sub SetPlayerAccess(ByVal index As Long, ByVal Access As Byte)
    Account(index).Chars(GetPlayerChar(index)).Access = Access
End Sub

Function GetPlayerPK(ByVal index As Long) As Byte

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Account(index).Chars(GetPlayerChar(index)).PK
End Function

Sub SetPlayerPK(ByVal index As Long, ByVal PK As Byte)
    Account(index).Chars(GetPlayerChar(index)).PK = PK
End Sub

Function GetPlayerVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Account(index).Chars(GetPlayerChar(index)).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Account(index).Chars(GetPlayerChar(index)).Vital(Vital) = Value

    If GetPlayerVital(index, Vital) > GetPlayerMaxVital(index, Vital) Then
        Account(index).Chars(GetPlayerChar(index)).Vital(Vital) = GetPlayerMaxVital(index, Vital)
    End If

    If GetPlayerVital(index, Vital) < 0 Then
        Account(index).Chars(GetPlayerChar(index)).Vital(Vital) = 0
    End If
End Sub

Function GetPlayerStat(ByVal index As Long, ByVal Stat As Stats) As Long
    Dim x As Long, i As Long
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    
    x = Account(index).Chars(GetPlayerChar(index)).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Account(index).Chars(GetPlayerChar(index)).Equipment(i).Num > 0 Then
            If Item(Account(index).Chars(GetPlayerChar(index)).Equipment(i).Num).Add_Stat(Stat) > 0 Then
                x = x + Item(Account(index).Chars(GetPlayerChar(index)).Equipment(i).Num).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = x
End Function

Function GetPlayerRawStat(ByVal index As Long, ByVal Stat As Stats) As Long
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerRawStat = Account(index).Chars(GetPlayerChar(index)).Stat(Stat)
End Function

Sub SetPlayerStat(ByVal index As Long, ByVal Stat As Stats, ByVal Value As Long, Optional ByVal PlusVal As Boolean = False)
    If Not PlusVal Then
        Account(index).Chars(GetPlayerChar(index)).Stat(Stat) = Value
    Else
        Account(index).Chars(GetPlayerChar(index)).Stat(Stat) = Account(index).Chars(GetPlayerChar(index)).Stat(Stat) + Value
    End If
End Sub

Function GetPlayerPoints(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerPoints = Account(index).Chars(GetPlayerChar(index)).Points
End Function

Sub SetPlayerPoints(ByVal index As Long, ByVal Points As Integer, Optional ByVal PlusVal As Boolean = False)
    If Not PlusVal Then
        Account(index).Chars(GetPlayerChar(index)).Points = Points
    Else
        Account(index).Chars(GetPlayerChar(index)).Points = Account(index).Chars(GetPlayerChar(index)).Points + Points
    End If
End Sub

Function GetPlayerMap(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Account(index).Chars(GetPlayerChar(index)).Map
End Function

Sub SetPlayerMap(ByVal index As Long, ByVal MapNum As Integer)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Account(index).Chars(GetPlayerChar(index)).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Account(index).Chars(GetPlayerChar(index)).x
End Function

Sub SetPlayerX(ByVal index As Long, ByVal x As Long)
    
    Account(index).Chars(GetPlayerChar(index)).x = x
End Sub

Function GetPlayerY(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Account(index).Chars(GetPlayerChar(index)).Y
End Function

Sub SetPlayerY(ByVal index As Long, ByVal Y As Long)
    Account(index).Chars(GetPlayerChar(index)).Y = Y
End Sub

Function GetPlayerDir(ByVal index As Long) As Long
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Account(index).Chars(GetPlayerChar(index)).Dir
End Function

Sub SetPlayerDir(ByVal index As Long, ByVal Dir As Byte)
    If Dir < DIR_UP Or Dir > DIR_DOWNRIGHT Then Exit Sub
    Account(index).Chars(GetPlayerChar(index)).Dir = Dir
End Sub

Function GetPlayerIP(ByVal index As Long) As String

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Byte) As Long
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Function
    
    GetPlayerInvItemNum = Account(index).Chars(GetPlayerChar(index)).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal index As Long, ByVal InvSlot As Byte, ByVal ItemNum As Integer)
    Account(index).Chars(GetPlayerChar(index)).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Byte) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Function
    GetPlayerInvItemValue = Account(index).Chars(GetPlayerChar(index)).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal index As Long, ByVal InvSlot As Byte, ByVal ItemValue As Long)
    Account(index).Chars(GetPlayerChar(index)).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Byte) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Account(index).Chars(GetPlayerChar(index)).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal index As Long, ByVal SpellSlot As Byte, ByVal SpellNum As Long)
    Account(index).Chars(GetPlayerChar(index)).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerSpellCD(ByVal index As Long, ByVal SpellSlot As Byte) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerSpellCD = Account(index).Chars(GetPlayerChar(index)).SpellCD(SpellSlot)
End Function

Sub SetPlayerSpellCD(ByVal index As Long, ByVal SpellSlot As Byte, ByVal NewCD As Long)
    Account(index).Chars(GetPlayerChar(index)).SpellCD(SpellSlot) = NewCD
End Sub

Function GetPlayerSwitch(ByVal index As Long, ByVal SwitchNum As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If SwitchNum < 1 Or SwitchNum > MAX_SWITCHES Then Exit Function
    GetPlayerSwitch = Account(index).Chars(GetPlayerChar(index)).Switches(SwitchNum)
End Function

Sub SetPlayerSwitch(ByVal index As Long, ByVal SwitchNum As Long, ByVal NewValue As Long)
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    If SwitchNum < 1 Or SwitchNum > MAX_SWITCHES Then Exit Sub
    Account(index).Chars(GetPlayerChar(index)).Switches(SwitchNum) = NewValue
End Sub

Function GetPlayerVariable(ByVal index As Long, ByVal VarNum As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If VarNum < 1 Or VarNum > MAX_VARIABLES Then Exit Function
    GetPlayerVariable = Account(index).Chars(GetPlayerChar(index)).Variables(VarNum)
End Function

Sub SetPlayerVariable(ByVal index As Long, ByVal VarNum As Long, ByVal NewValue As Long)
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    If VarNum < 1 Or VarNum > MAX_VARIABLES Then Exit Sub
    Account(index).Chars(GetPlayerChar(index)).Variables(VarNum) = NewValue
End Sub

Function GetPlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Byte) As Byte

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Or EquipmentSlot > Equipment.Equipment_Count - 1 Then Exit Function
    GetPlayerEquipment = Account(index).Chars(GetPlayerChar(index)).Equipment(EquipmentSlot).Num
End Function

Sub SetPlayerEquipment(ByVal index As Long, ByVal InvNum As Byte, ByVal EquipmentSlot As Byte)
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    Account(index).Chars(GetPlayerChar(index)).Equipment(EquipmentSlot).Num = InvNum
End Sub

Function GetPlayerEquipmentDur(ByVal index As Long, ByVal EquipmentSlot As Byte) As Integer

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Or EquipmentSlot > Equipment.Equipment_Count - 1 Then Exit Function
    GetPlayerEquipmentDur = Account(index).Chars(GetPlayerChar(index)).Equipment(EquipmentSlot).Durability
End Function

Sub SetPlayerEquipmentDur(ByVal index As Long, ByVal DurValue As Integer, ByVal EquipmentSlot As Byte)
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    Account(index).Chars(GetPlayerChar(index)).Equipment(EquipmentSlot).Durability = DurValue
End Sub

Function GetPlayerEquipmentBind(ByVal index As Long, ByVal EquipmentSlot As Byte) As Byte

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentBind = Account(index).Chars(GetPlayerChar(index)).Equipment(EquipmentSlot).Bind
End Function

Sub SetPlayerEquipmentBind(ByVal index As Long, ByVal BindType As Byte, ByVal EquipmentSlot As Byte)
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Sub
    Account(index).Chars(GetPlayerChar(index)).Equipment(EquipmentSlot).Bind = BindType
End Sub

Function GetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Byte) As Integer
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemNum = Account(index).Bank.Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal index As Long, ByVal BankSlot As Byte, ByVal ItemNum As Integer)
    Account(index).Bank.Item(BankSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Byte) As Long
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemValue = Account(index).Bank.Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal index As Long, ByVal BankSlot As Byte, ByVal ItemValue As Long)
    Account(index).Bank.Item(BankSlot).Value = ItemValue
End Sub

Function GetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Byte) As Long
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemDur = Account(index).Bank.Item(BankSlot).Durability
End Function

Sub SetPlayerBankItemDur(ByVal index As Long, ByVal BankSlot As Byte, ByVal DurValue As Long)
    Account(index).Bank.Item(BankSlot).Durability = DurValue
End Sub

Function GetPlayerBankItemBind(ByVal index As Long, ByVal BankSlot As Byte) As Long
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemBind = Account(index).Bank.Item(BankSlot).Bind
End Function

Sub SetPlayerBankItemBind(ByVal index As Long, ByVal BankSlot As Byte, ByVal BindValue As Long)
    Account(index).Bank.Item(BankSlot).Bind = BindValue
End Sub

Function GetPlayerGender(ByVal index As Long) As Long

    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerGender = Account(index).Chars(GetPlayerChar(index)).Gender
    Exit Function
End Function

Sub SetPlayerGender(ByVal index As Long, GenderNum As Byte)
    Account(index).Chars(GetPlayerChar(index)).Gender = GenderNum
End Sub

Function GetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Byte) As Integer
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Function
    GetPlayerInvItemDur = Account(index).Chars(GetPlayerChar(index)).Inv(InvSlot).Durability
End Function

Sub SetPlayerInvItemDur(ByVal index As Long, ByVal InvSlot As Byte, ByVal ItemDur As Integer)
    Account(index).Chars(GetPlayerChar(index)).Inv(InvSlot).Durability = ItemDur
End Sub

Function GetPlayerInvItemBind(ByVal index As Long, ByVal InvSlot As Byte) As Integer
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Function
    GetPlayerInvItemBind = Account(index).Chars(GetPlayerChar(index)).Inv(InvSlot).Bind
End Function

Sub SetPlayerInvItemBind(ByVal index As Long, ByVal InvSlot As Byte, ByVal BindType As Byte)
    Account(index).Chars(GetPlayerChar(index)).Inv(InvSlot).Bind = BindType
End Sub

Function GetMapItemX(ByVal MapNum As Integer, ByVal MapItemNum As Integer)
    GetMapItemX = MapItem(MapNum, MapItemNum).x
End Function

Sub SetMapItemX(ByVal MapNum As Integer, ByVal MapItemNum As Integer, ByVal Value As Long)
    MapItem(MapNum, MapItemNum).x = Value
End Sub

Function GetMapItemY(ByVal MapNum As Integer, ByVal MapItemNum As Integer)
    GetMapItemY = MapItem(MapNum, MapItemNum).Y
End Function

Sub SetMapItemY(ByVal MapNum As Integer, ByVal MapItemNum As Integer, ByVal Value As Long)
    MapItem(MapNum, MapItemNum).Y = Value
End Sub

Function GetPlayerHDSerial(ByVal index As Long) As String
    
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    GetPlayerHDSerial = Trim$(tempplayer(index).HDSerial)
End Function

Function GetClassName(ByVal ClassesNum As Long) As String
    GetClassName = Trim$(Class(ClassesNum).Name)
End Function

Function GetClasseStat(ByVal ClassesNum As Long, ByVal Stat As Stats) As Long
    GetClasseStat = Class(ClassesNum).Stat(Stat)
End Function

Function GetPlayerProficiency(ByVal index As Long, ByVal ProficiencyNum As Byte) As Long
    If index < 1 Or index > MAX_PLAYERS Then Exit Function
    
    Select Case Class(GetPlayerClass(index)).CombatTree
        Case 1: ' Melee
            If ProficiencyNum = Proficiency.Axe Or ProficiencyNum = Proficiency.Dagger Or ProficiencyNum = Proficiency.Mace Or ProficiencyNum = Proficiency.Spear Or ProficiencyNum = Proficiency.Sword Or ProficiencyNum = Proficiency.Heavy Or ProficiencyNum = Proficiency.Light Or ProficiencyNum = Proficiency.Medium Then
                GetPlayerProficiency = 1
            Else
                GetPlayerProficiency = 0
            End If
        Case 2: ' Range
            If ProficiencyNum = Proficiency.Dagger Or ProficiencyNum = Proficiency.Bow Or ProficiencyNum = Proficiency.Crossbow Or ProficiencyNum = Proficiency.Light Or ProficiencyNum = Proficiency.Medium Then
                GetPlayerProficiency = 1
            Else
                GetPlayerProficiency = 0
            End If
        Case 3: ' Magic
            If ProficiencyNum = Proficiency.Staff Or ProficiencyNum = Proficiency.Mace Or ProficiencyNum = Proficiency.Light Then
                GetPlayerProficiency = 1
            Else
                GetPlayerProficiency = 0
            End If
    End Select
End Function
