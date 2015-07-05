Attribute VB_Name = "modHandleData"
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

Sub LogPacket(Prse As String)
Dim FileName As String
Dim Hold As Integer

    FileName = App.Path & "\packet.log"
    Hold = Val(GetVar(FileName, "PACKETS", Prse))
    Call PutVar(FileName, "PACKETS", Prse, STR(Hold) + 1)
End Sub

Sub HandleData(ByVal Index As Long, ByVal Data As String)
'On Error Resume Next

Dim Parse() As String
Dim Name As String
Dim Password As String
Dim Sex As Long
Dim Class As Long
Dim CharNum As Long
Dim Msg As String
Dim MsgTo As String
Dim Dir As Long
Dim InvNum As Long
Dim Ammount As Long
Dim Damage As Long
Dim Movement As Long
Dim i As Long, n As Long, x As Long, y As Long, f As Long, z As Long
Dim MapNum As Long
Dim s As String
Dim ShopNum As Long, ItemNum As Long
Dim DurNeeded As Long, GoldNeeded As Long

    ' Handle Data
    Parse = Split(Data, SEP_CHAR)
    
    ' ::::::::::::::::::::::::
    ' :: New account packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "newaccount" Then
    Call LogPacket(Trim(Parse(0)))
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
        
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
            
            ' Check to see if account already exists
            If Not AccountExist(Name) Then
                Call AddAccount(Index, Name, Password)
                Call TextAdd(frmCServer.txtText, "Account " & Name & " has been created.", True)
                Call AddLog("Account " & Name & " has been created.", PLAYER_LOG)
                Call AlertMsg(Index, "Your account has been created!")
            Else
                Call AlertMsg(Index, "Sorry, that account name is already taken!")
            End If
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Delete account packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "delaccount" Then
    Call LogPacket(Trim(Parse(0)))
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
            
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(Index, "The name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If
            
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If
                        
            ' Delete names from master name file
            Call LoadPlayer(Index, Name)
            For i = 1 To MAX_CHARS
                If Trim(Player(Index).Char(i).Name) <> "" Then
                    Call DeleteName(Player(Index).Char(i).Name)
                End If
            Next i
            Call ClearPlayer(Index)
            
            ' Everything went ok
            Call Kill(App.Path & "\accounts\" & Trim(Name) & ".ini")
            Call TextAdd(frmCServer.txtText, "Account " & Name & " has been deleted.", True)
            Call AddLog("Account " & Trim(Name) & " has been deleted.", PLAYER_LOG)
            Call AlertMsg(Index, "Your account has been deleted.")
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::
    ' :: Login packet ::
    ' ::::::::::::::::::
    If LCase(Parse(0)) = "login" Then
    Call LogPacket(Trim(Parse(0)))
        If Not IsPlaying(Index) And Not IsLoggedIn(Index) Then
            ' Get the data
            Name = Parse(1)
            Password = Parse(2)
        
            ' Check versions
            If Val(Parse(3)) < CLIENT_MAJOR Or Val(Parse(4)) < CLIENT_MINOR Or Val(Parse(5)) < CLIENT_REVISION Then
                Call AlertMsg(Index, "Version outdated, please visit " & GAME_WEBSITE)
                Exit Sub
            End If
            
            If Len(Trim(Name)) < 3 Or Len(Trim(Password)) < 3 Then
                Call AlertMsg(Index, "Your name and password must be at least three characters in length")
                Exit Sub
            End If
            
            If Not AccountExist(Name) Then
                Call AlertMsg(Index, "That account name does not exist.")
                Exit Sub
            End If
        
            If Not PasswordOK(Name, Password) Then
                Call AlertMsg(Index, "Incorrect password.")
                Exit Sub
            End If
        
            'If IsMultiAccounts(Name) Then
                'Call AlertMsg(Index, "Multiple account logins is not authorized.")
                'Exit Sub
            'End If
                
            ' Everything went ok
            Dim Packs As String
            Packs = "MAXINFO" & SEP_CHAR
            Packs = Packs & GAME_NAME & SEP_CHAR
            Packs = Packs & GAME_WEBSITE & SEP_CHAR
            Packs = Packs & MAX_PLAYERS & SEP_CHAR
            Packs = Packs & MAX_ITEMS & SEP_CHAR
            Packs = Packs & MAX_NPCS & SEP_CHAR
            Packs = Packs & MAX_SHOPS & SEP_CHAR
            Packs = Packs & MAX_SPELLS & SEP_CHAR
            Packs = Packs & MAX_SKILLS & SEP_CHAR
            Packs = Packs & MAX_MAPS & SEP_CHAR
            Packs = Packs & MAX_MAP_ITEMS & SEP_CHAR
            Packs = Packs & MAX_GUILDS & SEP_CHAR
            Packs = Packs & MAX_GUILD_MEMBERS & SEP_CHAR
            'Packs = Packs & MAX_MAPX & SEP_CHAR
            'Packs = Packs & MAX_MAPY & SEP_CHAR
            'Packs = Packs & MAX_EMOTICONS & SEP_CHAR
            Packs = Packs & MAX_QUESTS & SEP_CHAR
            Packs = Packs & END_CHAR
            Call SendDataTo(Index, Packs)
    
            ' Load the player
            Call LoadPlayer(Index, Name)
            Call SendChars(Index)
    
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", PLAYER_LOG)
            Call TextAdd(frmCServer.txtText, GetPlayerLogin(Index) & " has logged in from " & GetPlayerIP(Index) & ".", True)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Requesting classes for making a character ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "getclasses" Then
    Call LogPacket(Trim(Parse(0)))
        If Not IsPlaying(Index) Then
            Call SendNewCharClasses(Index)
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Add character packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "addchar" Then
    Call LogPacket(Trim(Parse(0)))
        If Not IsPlaying(Index) Then
            Name = Parse(1)
            Sex = Val(Parse(2))
            Class = Val(Parse(3))
            CharNum = Val(Parse(4))
        
            ' Prevent hacking
            If Len(Trim(Name)) < 3 Then
                Call AlertMsg(Index, "Character name must be at least three characters in length.")
                Exit Sub
            End If
            
            ' Prevent being me
            If LCase(Trim(Name)) = "ambientiger" Then
                Call AlertMsg(Index, "Lets get one thing straight, you are not me, ok? :)")
                Exit Sub
            End If
            
            ' Prevent hacking
            For i = 1 To Len(Name)
                n = Asc(Mid(Name, i, 1))
                
                If (n >= 65 And n <= 90) Or (n >= 97 And n <= 122) Or (n = 95) Or (n = 32) Or (n >= 48 And n <= 57) Then
                Else
                    Call AlertMsg(Index, "Invalid name, only letters, numbers, spaces, and _ allowed in names.")
                    Exit Sub
                End If
            Next i
                                    
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Prevent hacking
            If (Sex < SEX_MALE) Or (Sex > SEX_FEMALE) Then
                Call HackingAttempt(Index, "Invalid Sex (dont laugh)")
                Exit Sub
            End If
            
            ' Prevent hacking
            If Class < 0 Or Class > Max_Classes Then
                Call HackingAttempt(Index, "Invalid Class")
                Exit Sub
            End If
        
            ' Check if char already exists in slot
            If CharExist(Index, CharNum) Then
                Call AlertMsg(Index, "Character already exists!" & vbNewLine & "Please Login again to continue.")
                Exit Sub
            End If
            
            ' Check if name is already in use
            If FindChar(Name) Then
                Call AlertMsg(Index, "Sorry, but that name is in use!" & vbNewLine & "Please Login again to continue.")
                Exit Sub
            End If
        
            ' Everything went ok, add the character
            Call AddChar(Index, Name, Sex, Class, CharNum)
            Call SavePlayer(Index)
            Call AddLog("Character " & Name & " added to " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(Index, "Character has been created!" & vbNewLine & "Please Login again to continue.")
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Deleting character packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "delchar" Then
    Call LogPacket(Trim(Parse(0)))
        If Not IsPlaying(Index) Then
            CharNum = Val(Parse(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
            
            Call DelChar(Index, CharNum)
            Call AddLog("Character deleted on " & GetPlayerLogin(Index) & "'s account.", PLAYER_LOG)
            Call AlertMsg(Index, "Character has been deleted!" & vbNewLine & "Please Login again to continue.")
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Using character packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "usechar" Then
    Call LogPacket(Trim(Parse(0)))
        If Not IsPlaying(Index) Then
            CharNum = Val(Parse(1))
        
            ' Prevent hacking
            If CharNum < 1 Or CharNum > MAX_CHARS Then
                Call HackingAttempt(Index, "Invalid CharNum")
                Exit Sub
            End If
        
            ' Check to make sure the character exists and if so, set it as its current char
            If CharExist(Index, CharNum) Then
                Player(Index).CharNum = CharNum
                Call JoinGame(Index)
            
                CharNum = Player(Index).CharNum
                Call AddLog(GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", PLAYER_LOG)
                Call TextAdd(frmCServer.txtText, GetPlayerLogin(Index) & "/" & GetPlayerName(Index) & " has began playing " & GAME_NAME & ".", True)
                Call UpdateCaption
                
                ' Now we want to check if they are already on the master list (this makes it add the user if they already haven't been added to the master list for older accounts)
                If Not FindChar(GetPlayerName(Index)) Then
                    f = FreeFile
                    Open App.Path & "\accounts\charlist.txt" For Append As #f
                        Print #f, GetPlayerName(Index)
                    Close #f
                End If
            Else
                Call AlertMsg(Index, "Character does not exist!")
            End If
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playermove" And Player(Index).GettingMap = NO Then
    Call LogPacket(Trim(Parse(0)))
        Dir = Val(Parse(1))
        Movement = Val(Parse(2))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(Index, "Invalid Direction")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Movement < 1 Or Movement > 2 Then
            Call HackingAttempt(Index, "Invalid Movement")
            Exit Sub
        End If
        
        ' Prevent player from moving if they have casted a spell
        If Player(Index).CastedSpell = YES Then
            ' Check if they have already casted a spell, and if so we can't let them move
            If GetTickCount > Player(Index).AttackTimer + 1000 Then
                Player(Index).CastedSpell = NO
            Else
                Call SendPlayerXY(Index)
                Exit Sub
            End If
        End If
        
        Call PlayerMove(Index, Dir, Movement)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Moving character packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "playerdir" And Player(Index).GettingMap = NO Then
    Call LogPacket(Trim(Parse(0)))
        Dir = Val(Parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(Index, "Invalid Direction")
            Exit Sub
        End If
        
        Call SetPlayerDir(Index, Dir)
        Call SendDataToMapBut(Index, GetPlayerMap(Index), "PLAYERDIR" & SEP_CHAR & Index & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::
    ' :: Social packets ::
    ' ::::::::::::::::::::
    If LCase(Parse(0)) = "saymsg" Then
    Call LogPacket(Trim(Parse(0)))
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Say Text Modification")
                Exit Sub
            End If
        Next i
        
        Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " says, '" & Msg & "'", PLAYER_LOG)
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " says, '" & Msg & "'", SayColor)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "emotemsg" Then
    Call LogPacket(Trim(Parse(0)))
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Emote Text Modification")
                Exit Sub
            End If
        Next i
        
        Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " " & Msg, PLAYER_LOG)
        Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " " & Msg, EmoteColor)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "broadcastmsg" Then
    Call LogPacket(Trim(Parse(0)))
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Broadcast Text Modification")
                Exit Sub
            End If
        Next i
        
        s = GetPlayerName(Index) & ": " & Msg
        Call AddLog(s, PLAYER_LOG)
        Call GlobalMsg(s, BroadcastColor)
        Call TextAdd(frmCServer.txtText, s, True)
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "globalmsg" Then
    Call LogPacket(Trim(Parse(0)))
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Global Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(Index) > 0 Then
            s = "(global) " & GetPlayerName(Index) & ": " & Msg
            Call AddLog(s, ADMIN_LOG)
            Call GlobalMsg(s, GlobalColor)
            Call TextAdd(frmCServer.txtText, s, True)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "adminmsg" Then
    Call LogPacket(Trim(Parse(0)))
        Msg = Parse(1)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Admin Text Modification")
                Exit Sub
            End If
        Next i
        
        If GetPlayerAccess(Index) > 0 Then
            Call AddLog("(admin " & GetPlayerName(Index) & ") " & Msg, ADMIN_LOG)
            Call AdminMsg("(admin " & GetPlayerName(Index) & ") " & Msg, AdminColor)
        End If
        Exit Sub
    End If
    
    If LCase(Parse(0)) = "playermsg" Then
    Call LogPacket(Trim(Parse(0)))
        MsgTo = FindPlayer(Parse(1))
        Msg = Parse(2)
        
        ' Prevent hacking
        For i = 1 To Len(Msg)
            If Asc(Mid(Msg, i, 1)) < 32 Or Asc(Mid(Msg, i, 1)) > 126 Then
                Call HackingAttempt(Index, "Player Msg Text Modification")
                Exit Sub
            End If
        Next i
        
        ' Check if they are trying to talk to themselves
        If MsgTo <> Index Then
            If MsgTo > 0 Then
                Call AddLog(GetPlayerName(Index) & " tells " & GetPlayerName(MsgTo) & ", " & Msg & "'", PLAYER_LOG)
                Call PlayerMsg(MsgTo, GetPlayerName(Index) & " tells you, '" & Msg & "'", TellColor)
                Call PlayerMsg(Index, "You tell " & GetPlayerName(MsgTo) & ", '" & Msg & "'", TellColor)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call AddLog("Map #" & GetPlayerMap(Index) & ": " & GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", PLAYER_LOG)
            Call MapMsg(GetPlayerMap(Index), GetPlayerName(Index) & " begins to mumble to himself, what a wierdo...", Green)
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Use item packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "useitem" Then
    Call LogPacket(Trim(Parse(0)))
        InvNum = Val(Parse(1))
        CharNum = Player(Index).CharNum
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If CharNum < 1 Or CharNum > MAX_CHARS Then
            Call HackingAttempt(Index, "Invalid CharNum")
            Exit Sub
        End If
        
        If (GetPlayerInvItemNum(Index, InvNum) > 0) And (GetPlayerInvItemNum(Index, InvNum) <= MAX_ITEMS) Then
            n = Item(GetPlayerInvItemNum(Index, InvNum)).Data2
            
            ' Find out what kind of item it is
            Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Type
                Case ITEM_TYPE_ARMOR
                    If InvNum <> GetPlayerArmorSlot(Index) Then
                        If Int(GetPlayerDEF(Index)) < n Then
                            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Defense Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            Exit Sub
                        End If
                        Call SetPlayerArmorSlot(Index, InvNum)
                    Else
                        Call SetPlayerArmorSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                
                Case ITEM_TYPE_WEAPON
                    If InvNum <> GetPlayerWeaponSlot(Index) Then
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Data3 >= WEAPON_SUBTYPE_DAGGER And Item(GetPlayerInvItemNum(Index, InvNum)).Data3 <= WEAPON_SUBTYPE_MACE Then
                            If Int(GetPlayerSTR(Index)) < n Then
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Strength Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                Exit Sub
                            End If
                        ElseIf Item(GetPlayerInvItemNum(Index, InvNum)).Data3 >= WEAPON_SUBTYPE_WAND And Item(GetPlayerInvItemNum(Index, InvNum)).Data3 <= WEAPON_SUBTYPE_STAFF Then
                            If Int(GetPlayerMAGI(Index)) < n Then
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Magic Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                Exit Sub
                            End If
                        Else
                            If Int(GetPlayerSTR(Index)) < n Then
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Strength Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                Exit Sub
                            End If
                        End If
                        Call SetPlayerWeaponSlot(Index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                        
                Case ITEM_TYPE_HELMET
                    If InvNum <> GetPlayerHelmetSlot(Index) Then
                        If Int(GetPlayerDEF(Index)) < n Then
                            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Defence Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            Exit Sub
                        End If
                        Call SetPlayerHelmetSlot(Index, InvNum)
                    Else
                        Call SetPlayerHelmetSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
            
                Case ITEM_TYPE_SHIELD
                    If InvNum <> GetPlayerShieldSlot(Index) Then
                        If Int(GetPlayerDEF(Index)) < n Then
                            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Defence Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            Exit Sub
                        End If
                        Call SetPlayerShieldSlot(Index, InvNum)
                    Else
                        Call SetPlayerShieldSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
            
                Case ITEM_TYPE_POTIONADDHP
                    Call SetPlayerHP(Index, GetPlayerHP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "HP Gained: +" & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(Index)
                
                Case ITEM_TYPE_POTIONADDMP
                    Call SetPlayerMP(Index, GetPlayerMP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "MP Gained: +" & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(Index)
        
                Case ITEM_TYPE_POTIONADDSP
                    Call SetPlayerSP(Index, GetPlayerSP(Index) + Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "SP Gained: +" & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & Green & SEP_CHAR & END_CHAR)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(Index)

                Case ITEM_TYPE_POTIONSUBHP
                    Call SetPlayerHP(Index, GetPlayerHP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "HP Lost: +" & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendHP(Index)
                
                Case ITEM_TYPE_POTIONSUBMP
                    Call SetPlayerMP(Index, GetPlayerMP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "MP Lost: +" & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendMP(Index)
        
                Case ITEM_TYPE_POTIONSUBSP
                    Call SetPlayerSP(Index, GetPlayerSP(Index) - Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1)
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "SP Lost: +" & Item(Player(Index).Char(CharNum).Inv(InvNum).Num).Data1 & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Call TakeItem(Index, Player(Index).Char(CharNum).Inv(InvNum).Num, 0)
                    Call SendSP(Index)
                    
                Case ITEM_TYPE_KEY
                    Select Case GetPlayerDir(Index)
                        Case DIR_UP
                            If GetPlayerY(Index) > 0 Then
                                x = GetPlayerX(Index)
                                y = GetPlayerY(Index) - 1
                            Else
                                Exit Sub
                            End If
                            
                        Case DIR_DOWN
                            If GetPlayerY(Index) < MAX_MAPY Then
                                x = GetPlayerX(Index)
                                y = GetPlayerY(Index) + 1
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_LEFT
                            If GetPlayerX(Index) > 0 Then
                                x = GetPlayerX(Index) - 1
                                y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                                
                        Case DIR_RIGHT
                            If GetPlayerX(Index) < MAX_MAPY Then
                                x = GetPlayerX(Index) + 1
                                y = GetPlayerY(Index)
                            Else
                                Exit Sub
                            End If
                    End Select
                    
                    ' Check if a key exists
                    If Map(GetPlayerMap(Index)).Tile(x, y).Type = TILE_TYPE_KEY Then
                        ' Check if the key they are using matches the map key
                        If GetPlayerInvItemNum(Index, InvNum) = Map(GetPlayerMap(Index)).Tile(x, y).Data1 Then
                            TempTile(GetPlayerMap(Index)).DoorOpen(x, y) = YES
                            TempTile(GetPlayerMap(Index)).DoorTimer = GetTickCount
                            
                            Call SendDataToMap(GetPlayerMap(Index), "MAPKEY" & SEP_CHAR & x & SEP_CHAR & y & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
                            'Call MapMsg(GetPlayerMap(Index), "A door has been unlocked.", White)
                            
                            ' Check if we are supposed to take away the item
                            If Map(GetPlayerMap(Index)).Tile(x, y).Data2 = 1 Then
                                Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Key Disolves" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            End If
                        End If
                    End If
                    
                Case ITEM_TYPE_SPELL
                    ' Get the spell num
                    n = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If Spell(n).ClassReq - 1 = GetPlayerClass(Index) Or Spell(n).ClassReq = 0 Then
                            ' Make sure they are the right level
                            i = GetSpellReqLevel(Index, n)
                            If i <= GetPlayerLevel(Index) Then
                                i = FindOpenSpellSlot(Index)
                                
                                ' Make sure they have an open spell slot
                                If i > 0 Then
                                    ' Make sure they dont already have the spell
                                    If Not HasSpell(Index, n) Then
                                        Call SetPlayerSpell(Index, i, n)
                                        Call SetPlayerSpellLevel(Index, i, 1)
                                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                        Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Studying Spell...." & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "New Spell Learned" & SEP_CHAR & DarkGrey & SEP_CHAR & END_CHAR)
                                        Call SendPlayerSpells(Index)
                                    Else
                                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Spell Already Learned" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                    End If
                                Else
                                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "No Room For Spell" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                End If
                           Else
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Level Req: " & i & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            End If
                        Else
                            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Class Req: " & GetClassName(Spell(n).ClassReq - 1) & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Spell Not Connected" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Contact Admin" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    End If
                    
                Case ITEM_TYPE_TOOL
                    If InvNum <> GetPlayerWeaponSlot(Index) Then
                        If Item(GetPlayerInvItemNum(Index, InvNum)).Data3 >= WEAPON_SUBTYPE_DAGGER And Item(GetPlayerInvItemNum(Index, InvNum)).Data3 <= WEAPON_SUBTYPE_MACE Then
                            If Int(GetPlayerSTR(Index)) < n Then
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Strength Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                Exit Sub
                            End If
                        ElseIf Item(GetPlayerInvItemNum(Index, InvNum)).Data3 >= WEAPON_SUBTYPE_WAND And Item(GetPlayerInvItemNum(Index, InvNum)).Data3 <= WEAPON_SUBTYPE_STAFF Then
                            If Int(GetPlayerMAGI(Index)) < n Then
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Magic Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                Exit Sub
                            End If
                        Else
                            If Int(GetPlayerSTR(Index)) < n Then
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Strength Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                Exit Sub
                            End If
                        End If
                        Call SetPlayerWeaponSlot(Index, InvNum)
                    Else
                        Call SetPlayerWeaponSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
                    
                Case ITEM_TYPE_AMULET
                    If InvNum <> GetPlayerAmuletSlot(Index) Then
                        Call SetPlayerAmuletSlot(Index, InvNum)
                        Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                            Case CHARM_TYPE_ADDHP
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDMP
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSP
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSTR
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDDEF
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDMAGI
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSPEED
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDDEX
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDCRIT
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDDROP
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDBLOCK
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDACCU
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                        End Select
                        'If Int(GetPlayerSTR(Index)) < n Then
                            'Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Strength Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            'Exit Sub
                        'End If
                    Else
                        Call SetPlayerAmuletSlot(Index, 0)
                        Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                            Case CHARM_TYPE_ADDHP
                                Call SendHP(Index)
                                
                            Case CHARM_TYPE_ADDMP
                                Call SendMP(Index)
                                
                            Case CHARM_TYPE_ADDSP
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSTR
                                Call SendStats(Index)
                                Call SendHP(Index)
                                
                            Case CHARM_TYPE_ADDDEF
                                Call SendStats(Index)
                                
                            Case CHARM_TYPE_ADDMAGI
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSPEED
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDDEX
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                        End Select
                    End If
                    Call SendWornEquipment(Index)
                    
                Case ITEM_TYPE_RING
                    If InvNum <> GetPlayerRingSlot(Index) Then
                        Call SetPlayerRingSlot(Index, InvNum)
                        Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                            Case CHARM_TYPE_ADDHP
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDMP
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSP
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSTR
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDDEF
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDMAGI
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSPEED
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDCRIT
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDDROP
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDBLOCK
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDACCU
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                        End Select
                        'If Int(GetPlayerSTR(Index)) < n Then
                            'Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Strength Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            'Exit Sub
                        'End If
                    Else
                        Call SetPlayerRingSlot(Index, 0)
                        Select Case Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                            Case CHARM_TYPE_ADDHP
                                Call SendHP(Index)
                                
                            Case CHARM_TYPE_ADDMP
                                Call SendMP(Index)
                                
                            Case CHARM_TYPE_ADDSP
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSTR
                                Call SendStats(Index)
                                Call SendHP(Index)
                                
                            Case CHARM_TYPE_ADDDEF
                                Call SendStats(Index)
                                
                            Case CHARM_TYPE_ADDMAGI
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDSPEED
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                                
                            Case CHARM_TYPE_ADDACCU
                                Call SendStats(Index)
                                Call SendHP(Index)
                                Call SendMP(Index)
                                Call SendSP(Index)
                        End Select
                    End If
                    Call SendWornEquipment(Index)
                                
                Case ITEM_TYPE_SKILL
                    ' Get the skill num
                    n = Item(GetPlayerInvItemNum(Index, InvNum)).Data1
                    
                    If n > 0 Then
                        ' Make sure they are the right class
                        If Skill(n).ClassReq - 1 = GetPlayerClass(Index) Or Skill(n).ClassReq = 0 Then
                            ' Make sure they are the right level
                            i = GetSkillReqLevel(Index, n)
                            If i <= GetPlayerLevel(Index) Then
                                i = FindOpenSkillSlot(Index)
                                
                                ' Make sure they have an open spell slot
                                If i > 0 Then
                                    ' Make sure they dont already have the skill
                                    If Not HasSkill(Index, n) Then
                                        Call SetPlayerSkill(Index, i, n)
                                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                        Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Studying Skill...." & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "New Skill Learned" & SEP_CHAR & DarkGrey & SEP_CHAR & END_CHAR)
                                        Call SendPlayerSkills(Index)
                                    Else
                                        Call TakeItem(Index, GetPlayerInvItemNum(Index, InvNum), 0)
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Skill Already Learned" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                                    End If
                                Else
                                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "No Room For Skill" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                                End If
                            Else
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Level Req: " & i & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            End If
                        Else
                            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Class Req: " & GetClassName(Skill(n).ClassReq - 1) & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Skill Not Connected" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Contact Admin" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    End If
                    
                    Case ITEM_TYPE_ARROW
                    If InvNum <> GetPlayerArrowSlot(Index) Then
                        If Int(GetPlayerDEX(Index)) < n Then
                            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Dexterity Too Low" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                            Exit Sub
                        End If
                        Call SetPlayerArrowSlot(Index, InvNum)
                    Else
                        Call SetPlayerArrowSlot(Index, 0)
                    End If
                    Call SendWornEquipment(Index)
            End Select
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Player attack packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "attack" Then
    Call LogPacket(Trim(Parse(0)))
        If GetPlayerWeaponSlot(Index) > 0 Then
            If Item(GetPlayerInvItemNum(Index, GetPlayerWeaponSlot(Index))).Data3 = WEAPON_SUBTYPE_BOW Then
                If GetPlayerArrowSlot(Index) > 0 Then
                    Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & Item(GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index))).Data3 & SEP_CHAR & GetPlayerDir(Index) & SEP_CHAR & END_CHAR)
                    Call SetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index), GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) - 1)
                    Call SetPlayerInvItemDur(Index, GetPlayerWeaponSlot(Index), GetPlayerInvItemDur(Index, GetPlayerWeaponSlot(Index)) - 1)
                    Call SendInventoryUpdate(Index, GetPlayerArrowSlot(Index))
                    Call SendInventoryUpdate(Index, GetPlayerWeaponSlot(Index))
                    Call AttributeSkillsExp(Index, SKILL_TYPE_ATTRIBUTE, SKILL_ATTRIBUTE_STR, 1, True)
                    'Call AttributeSkillsExp(Index, SKILL_TYPE_ATTRIBUTE, SKILL_ATTRIBUTE_DEX, 1, False)
                    If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) <= 0 Then
                        Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index))).Name) & " 's Depleted" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                        'Call TakeItem(Index, GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index)), 0)
                    Else
                        If GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) < 5 Then
                            Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & Trim(Item(GetPlayerInvItemNum(Index, GetPlayerArrowSlot(Index))).Name) & "'s - " & GetPlayerInvItemDur(Index, GetPlayerArrowSlot(Index)) & " Remaining!" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                        End If
                    End If
                    Exit Sub
                Else
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Equip Arrows" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
        End If
        ' Try to attack a player
        For i = 1 To HighIndex
            ' Make sure we dont try to attack ourselves
            If i <> Index Then
                ' Can we attack the player?
                If CanAttackPlayer(Index, i) Then
                    If Not CanPlayerBlockHit(i) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            Damage = GetPlayerDamage(Index) - GetPlayerProtection(i)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(i)
                            Call SendDataTo(Index, "BLITPKMSG" & SEP_CHAR & i & SEP_CHAR & "Critical Hit" & SEP_CHAR & BrightCyan & SEP_CHAR & END_CHAR)
                            Call SendDataTo(i, "BLITWARNMSG" & SEP_CHAR & "Critical Hit" & SEP_CHAR & BrightCyan & SEP_CHAR & END_CHAR)
                        End If
                        
                        If Damage > 0 Then
                            Call AttackPlayer(Index, i, Damage)
                        Else
                            Call SendDataTo(Index, "BLITPKMSG" & SEP_CHAR & "  Miss" & SEP_CHAR & i & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                            Call SendDataTo(i, "BLITPLAYERMSG" & SEP_CHAR & "  Miss" & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call SendDataTo(Index, "BLITPKMSG" & SEP_CHAR & "  Block" & SEP_CHAR & i & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                        Call SendDataTo(i, "BLITPLAYERMSG" & SEP_CHAR & "  Block" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                    End If
                    
                    Exit Sub
                End If
            End If
        Next i
        
        ' Try to attack a npc
        For i = 1 To MAX_MAP_NPCS
            ' Can we attack the npc?
            If CanAttackNpc(Index, i) Then
                ' Get the damage we can do
                If Not CanPlayerCriticalHit(Index) Then
                    Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                Else
                    n = GetPlayerDamage(Index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), i).Num).DEF / 2)
                    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & BrightCyan & SEP_CHAR & END_CHAR)
                End If
                
                If Damage > 0 Then
                    Call AttackNpc(Index, i, Damage)
                    Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & i & SEP_CHAR & White & SEP_CHAR & END_CHAR)
                Else
                    Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "  Miss" & SEP_CHAR & i & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
            End If
        Next i
        
        ' Try to attack a resource
        For i = 1 To MAX_MAP_RESOURCES
            ' Can we attack the resource?
            If CanAttackResource(Index, i) Then
                ' Get the damage we can do
                If Not CanPlayerCriticalHit(Index) Then
                    Damage = GetPlayerDamage(Index) - Int(Npc(MapResource(GetPlayerMap(Index), i).Num).DEF / 2)
                Else
                    n = GetPlayerDamage(Index)
                    Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapResource(GetPlayerMap(Index), i).Num).DEF / 2)
                    Call SendDataTo(Index, "BLITRESOURCEDMG" & SEP_CHAR & i & SEP_CHAR & Damage & SEP_CHAR & BrightCyan & SEP_CHAR & END_CHAR)
                End If
                
                If Damage > 0 Then
                    Call AttackResource(Index, i, Damage)
                    Call SendDataTo(Index, "BLITRESOURCEDMG" & SEP_CHAR & i & SEP_CHAR & Damage & SEP_CHAR & White & SEP_CHAR & END_CHAR)
                Else
                    Call SendDataTo(Index, "BLITRESOURCEMSG" & SEP_CHAR & i & SEP_CHAR & "  Miss" & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                End If
                Exit Sub
            End If
        Next i
        
        Exit Sub
    End If
    
    ':::::::::::::::::::
    ':: Arrow packets ::
    ':::::::::::::::::::
    'If LCase(Parse(0)) = "checkarrows" Then
        'n = Arrows(Val(Parse(1))).Pic
       
        'Call SendDataToMap(GetPlayerMap(Index), "checkarrows" & SEP_CHAR & Index & SEP_CHAR & n & SEP_CHAR & END_CHAR)
        'Exit Sub
    'End If
    
    If LCase(Parse(0)) = "lastarrow" Then
    Call LogPacket(Trim(Parse(0)))
        n = Val(Parse(1))
        
        Call TakeItem(Index, n, 0)
        Exit Sub
    End If
   
    If LCase(Parse(0)) = "arrowhit" Then
    Call LogPacket(Trim(Parse(0)))
        n = Val(Parse(1))
        z = Val(Parse(2))
        x = Val(Parse(3))
        y = Val(Parse(4))
       
        If n = TARGET_TYPE_PLAYER Then
            ' Make sure we dont try to attack ourselves
            If z <> Index Then
                ' Can we attack hit player?
                If IsAccurate(Index) Then
                    If Not CanPlayerBlockHit(z) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            Damage = GetPlayerDamage(Index) - GetPlayerProtection(z)
                            'Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - GetPlayerProtection(z)
                            Call SendDataTo(Index, "BLITPKMSG" & SEP_CHAR & z & SEP_CHAR & "Critical Hit" & SEP_CHAR & BrightCyan & SEP_CHAR & END_CHAR)
                            Call SendDataTo(z, "BLITWARNMSG" & SEP_CHAR & "Critical Hit" & SEP_CHAR & BrightCyan & SEP_CHAR & END_CHAR)
                       End If
                          
                        If Damage > 0 Then
                            Call AttackPlayer(Index, z, Damage)
                        Else
                            Call SendDataTo(Index, "BLITPKMSG" & SEP_CHAR & "No Damage" & SEP_CHAR & z & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                            Call SendDataTo(z, "BLITPLAYERMSG" & SEP_CHAR & "No Damage" & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call SendDataTo(Index, "BLITPKMSG" & SEP_CHAR & "  Block" & SEP_CHAR & z & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                        Call SendDataTo(i, "BLITPLAYERMSG" & SEP_CHAR & "  Block" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                    End If
                    Exit Sub
                Else
                    Call SendDataTo(Index, "BLITPKMSG" & SEP_CHAR & "  Miss" & SEP_CHAR & z & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                    Call SendDataTo(z, "BLITPLAYERMSG" & SEP_CHAR & "  Miss" & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                End If
            End If
        ElseIf n = TARGET_TYPE_NPC Then
            ' Check for friendly/shopkeeper npc
            If Npc(MapNpc(GetPlayerMap(Index), z).Num).Behavior <> NPC_BEHAVIOR_FRIENDLY Then
                If Npc(MapNpc(GetPlayerMap(Index), z).Num).Behavior <> NPC_BEHAVIOR_SHOPKEEPER Then
                    ' Can we hit the npc?
                    If IsAccurate(Index) Then
                        ' Get the damage we can do
                        If Not CanPlayerCriticalHit(Index) Then
                            Damage = GetPlayerDamage(Index) - Int(Npc(MapNpc(GetPlayerMap(Index), z).Num).DEF / 2)
                            'Call SendDataToMap(GetPlayerMap(Index), "sound" & SEP_CHAR & "attack" & SEP_CHAR & END_CHAR)
                        Else
                            n = GetPlayerDamage(Index)
                            Damage = n + Int(Rnd * Int(n / 2)) + 1 - Int(Npc(MapNpc(GetPlayerMap(Index), z).Num).DEF / 2)
                            Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Critical Hit" & SEP_CHAR & BrightCyan & SEP_CHAR & END_CHAR)
                        End If
                
                        If Damage > 0 Then
                            Call AttackNpc(Index, z, Damage)
                            Call SendDataTo(Index, "BLITPLAYERDMG" & SEP_CHAR & Damage & SEP_CHAR & z & SEP_CHAR & White & SEP_CHAR & END_CHAR)
                        Else
                            Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "No Damage" & SEP_CHAR & z & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                        End If
                        Exit Sub
                    Else
                        Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "  Miss" & SEP_CHAR & z & SEP_CHAR & Magenta & SEP_CHAR & END_CHAR)
                    End If
                Else
                    Call SendTrade(Index, Npc(MapNpc(GetPlayerMap(Index), z).Num).ShopLink)
                    Exit Sub
                End If
            Else
                Call SendNpcQuests(Index, MapNpc(GetPlayerMap(Index), z).Num)
                Exit Sub
            End If
        End If
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::
    ' :: Player request for a new map ::
    ' ::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requestnewmap" Then
    Call LogPacket(Trim(Parse(0)))
        Dir = Val(Parse(1))
        
        ' Prevent hacking
        If Dir < DIR_UP Or Dir > DIR_RIGHT Then
            Call HackingAttempt(Index, "Invalid Direction")
            Exit Sub
        End If
                
        Call PlayerMove(Index, Dir, 1)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Map data packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "mapdata" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = 1
        
        MapNum = GetPlayerMap(Index)
        Map(MapNum).Name = Parse(n + 1)
        Map(MapNum).Revision = Map(MapNum).Revision + 1
        Map(MapNum).Owner = Parse(n + 3)
        Map(MapNum).Moral = Val(Parse(n + 4))
        Map(MapNum).Up = Val(Parse(n + 5))
        Map(MapNum).Down = Val(Parse(n + 6))
        Map(MapNum).Left = Val(Parse(n + 7))
        Map(MapNum).Right = Val(Parse(n + 8))
        Map(MapNum).Music = Val(Parse(n + 9))
        Map(MapNum).BootMap = Val(Parse(n + 10))
        Map(MapNum).BootX = Val(Parse(n + 11))
        Map(MapNum).BootY = Val(Parse(n + 12))
        Map(MapNum).Indoors = Val(Parse(n + 13))
        
        n = n + 14
        
        For x = 1 To MAX_MAP_NPCS
            Map(MapNum).NSpawn(x).NSx = Val(Parse(n))
            Map(MapNum).NSpawn(x).NSy = Val(Parse(n + 1))
            
            n = n + 2
        Next x
        
        For x = 1 To MAX_MAP_RESOURCES
            Map(MapNum).RSpawn(x).RSx = Val(Parse(n))
            Map(MapNum).RSpawn(x).RSy = Val(Parse(n + 1))
            
            n = n + 2
        Next x
        
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                Map(MapNum).Tile(x, y).Tileset = Val(Parse(n))
                Map(MapNum).Tile(x, y).Ground = Val(Parse(n + 1))
                Map(MapNum).Tile(x, y).Mask = Val(Parse(n + 2))
                Map(MapNum).Tile(x, y).Mask2 = Val(Parse(n + 3))
                Map(MapNum).Tile(x, y).Anim = Val(Parse(n + 4))
                Map(MapNum).Tile(x, y).Fringe = Val(Parse(n + 5))
                Map(MapNum).Tile(x, y).Fringe2 = Val(Parse(n + 6))
                Map(MapNum).Tile(x, y).FAnim = Val(Parse(n + 7))
                Map(MapNum).Tile(x, y).Light = Val(Parse(n + 8))
                Map(MapNum).Tile(x, y).Type = Val(Parse(n + 9))
                Map(MapNum).Tile(x, y).Data1 = Val(Parse(n + 10))
                Map(MapNum).Tile(x, y).Data2 = Val(Parse(n + 11))
                Map(MapNum).Tile(x, y).Data3 = Val(Parse(n + 12))
                Map(MapNum).Tile(x, y).WalkUp = Parse(n + 13)
                Map(MapNum).Tile(x, y).WalkDown = Parse(n + 14)
                Map(MapNum).Tile(x, y).WalkLeft = Parse(n + 15)
                Map(MapNum).Tile(x, y).WalkRight = Parse(n + 16)
                Map(MapNum).Tile(x, y).Build = Parse(n + 17)
                
                n = n + 18
            Next x
        Next y
        
        For x = 1 To MAX_MAP_NPCS
            Map(MapNum).Npc(x) = Val(Parse(n))
            n = n + 1
            Call ClearMapNpc(x, MapNum)
        Next x
        For x = 1 To MAX_MAP_RESOURCES
            Map(MapNum).Resource(x) = Val(Parse(n))
            n = n + 1
            Call ClearMapResource(x, MapNum)
        Next x
        Call SendMapNpcsToMap(MapNum)
        Call SpawnMapNpcs(MapNum)
        Call SendMapResourcesToMap(MapNum)
        Call SpawnMapResources(MapNum)
        
        ' Clear it all out
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
            Call ClearMapItem(i, GetPlayerMap(Index))
        Next i
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(Index))
        
        ' Save the map
        Call SaveMap(MapNum)
        
        ' Refresh map for everyone online
        For i = 1 To HighIndex
            If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, "CHECKFORMAP" & SEP_CHAR & GetPlayerMap(i) & SEP_CHAR & Map(GetPlayerMap(i)).Revision & SEP_CHAR & END_CHAR)
            End If
        Next i
        
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::
    ' :: Need map yes/no packet ::
    ' ::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "needmap" Then
    Call LogPacket(Trim(Parse(0)))
        ' Get yes/no value
        s = LCase(Parse(1))
                
        If s = "yes" Then
            Call SendMap(Index, GetPlayerMap(Index))
            Call SendMapItemsTo(Index, GetPlayerMap(Index))
            Call SendMapNpcsTo(Index, GetPlayerMap(Index))
            Call SendMapResourcesTo(Index, GetPlayerMap(Index))
            Call SendJoinMap(Index)
            Player(Index).GettingMap = NO
            Call SendDataTo(Index, "MAPDONE" & SEP_CHAR & END_CHAR)
        Else
            Call SendMapItemsTo(Index, GetPlayerMap(Index))
            Call SendMapNpcsTo(Index, GetPlayerMap(Index))
            Call SendMapResourcesTo(Index, GetPlayerMap(Index))
            Call SendJoinMap(Index)
            Player(Index).GettingMap = NO
            Call SendDataTo(Index, "MAPDONE" & SEP_CHAR & END_CHAR)
        End If
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to pick up something packet ::
    ' :::::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapgetitem" Then
    Call LogPacket(Trim(Parse(0)))
        Call PlayerMapGetItem(Index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::::::::::::::::
    ' :: Player trying to drop something packet ::
    ' ::::::::::::::::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "mapdropitem" Then
    Call LogPacket(Trim(Parse(0)))
        InvNum = Val(Parse(1))
        Ammount = Val(Parse(2))
        
        ' Prevent hacking
        If InvNum < 1 Or InvNum > MAX_INV Then
            Call HackingAttempt(Index, "Invalid InvNum")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Ammount > GetPlayerInvItemValue(Index, InvNum) Then
            Call HackingAttempt(Index, "Item ammount modification")
            Exit Sub
        End If
        
        ' Prevent hacking
        If Item(GetPlayerInvItemNum(Index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
            ' Check if money and if it is we want to make sure that they aren't trying to drop 0 value
            If Ammount <= 0 Then
                Call HackingAttempt(Index, "Trying to drop 0 ammount of currency")
                Exit Sub
            End If
        End If
            
        Call PlayerMapDropItem(Index, InvNum, Ammount)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::
    ' :: Respawn map packet ::
    ' ::::::::::::::::::::::::
    If LCase(Parse(0)) = "maprespawn" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Clear out it all
        For i = 1 To MAX_MAP_ITEMS
            Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(Index), MapItem(GetPlayerMap(Index), i).x, MapItem(GetPlayerMap(Index), i).y)
            Call ClearMapItem(i, GetPlayerMap(Index))
        Next i
        
        ' Respawn
        Call SpawnMapItems(GetPlayerMap(Index))
        
        ' Respawn NPCS
        For i = 1 To MAX_MAP_NPCS
            Call SpawnNpc(i, GetPlayerMap(Index))
            Call SpawnResource(i, GetPlayerMap(Index))
        Next i
        
        Call PlayerMsg(Index, "Map respawned.", Blue)
        Call AddLog(GetPlayerName(Index) & " has respawned map #" & GetPlayerMap(Index), ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit map packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditmap" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_MAPPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "EDITMAP" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit item packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requestedititem" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "ITEMEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If

    ' ::::::::::::::::::::::
    ' :: Edit item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "edititem" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The item #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Item Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing item #" & n & ".", ADMIN_LOG)
        Call SendEditItemTo(Index, n)
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save item packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "saveitem" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        If n < 0 Or n > MAX_ITEMS Then
            Call HackingAttempt(Index, "Invalid Item Index")
            Exit Sub
        End If
        
        ' Update the item
        Item(n).Name = Parse(2)
        Item(n).Pic = Val(Parse(3))
        Item(n).Type = Val(Parse(4))
        Item(n).Data1 = Val(Parse(5))
        Item(n).Data2 = Val(Parse(6))
        Item(n).Data3 = Val(Parse(7))
        
        ' Save it
        Call SendUpdateItemToAll(n)
        Call SaveItem(n)
        Call AddLog(GetPlayerName(Index) & " saved item #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::
    ' :: Request edit npc packet ::
    ' :::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditnpc" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "NPCEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Edit npc packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "editnpc" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The npc #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(Index, "Invalid NPC Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing npc #" & n & ".", ADMIN_LOG)
        Call SendEditNpcTo(Index, n)
    End If
    
    ' :::::::::::::::::::::
    ' :: Save npc packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "savenpc" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_NPCS Then
            Call HackingAttempt(Index, "Invalid NPC Index")
            Exit Sub
        End If
        
        f = 1
        
        ' Update the npc
        Npc(n).Name = Parse(f + 1)
        Npc(n).Sprite = Val(Parse(f + 2))
        Npc(n).SpawnSecs = Val(Parse(f + 3))
        Npc(n).Behavior = Val(Parse(f + 4))
        Npc(n).Range = Val(Parse(f + 5))
        Npc(n).STR = Val(Parse(f + 6))
        Npc(n).DEF = Val(Parse(f + 7))
        Npc(n).SPEED = Val(Parse(f + 8))
        Npc(n).MAGI = Val(Parse(f + 9))
        Npc(n).Big = Val(Parse(f + 10))
        Npc(n).MaxHp = Val(Parse(f + 11))
        Npc(n).Respawn = Val(Parse(f + 12))
        Npc(n).HitOnlyWith = Val(Parse(f + 13))
        Npc(n).ShopLink = Val(Parse(f + 14))
        Npc(n).ExpType = Val(Parse(f + 15))
        Npc(n).EXP = Val(Parse(f + 16))
        
        f = f + 17
        
        For i = 1 To MAX_NPC_QUESTS
            Npc(n).QuestNPC(i) = Val(Parse((f - 1) + i))
        Next i
        
        f = f + MAX_NPC_QUESTS
        
        For i = 1 To MAX_NPC_DROPS
            Npc(n).ItemNPC(i).Chance = Val(Parse(f))
            Npc(n).ItemNPC(i).ItemNum = Val(Parse(f + 1))
            Npc(n).ItemNPC(i).ItemValue = Val(Parse(f + 2))
            
            f = f + 3
        Next i
        
        ' Save it
        Call SendUpdateNpcToAll(n)
        Call SaveNpc(n)
        Call AddLog(GetPlayerName(Index) & " saved npc #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::::::
    ' :: Request edit shop packet ::
    ' ::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditshop" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SHOPEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::
    ' :: Edit shop packet ::
    ' ::::::::::::::::::::::
    If LCase(Parse(0)) = "editshop" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The shop #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SHOPS Then
            Call HackingAttempt(Index, "Invalid Shop Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing shop #" & n & ".", ADMIN_LOG)
        Call SendEditShopTo(Index, n)
    End If
    
    ' ::::::::::::::::::::::
    ' :: Save shop packet ::
    ' ::::::::::::::::::::::
    If (LCase(Parse(0)) = "saveshop") Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ShopNum = Val(Parse(1))
        
        ' Prevent hacking
        If ShopNum < 0 Or ShopNum > MAX_SHOPS Then
            Call HackingAttempt(Index, "Invalid Shop Index")
            Exit Sub
        End If
        
        ' Update the shop
        Shop(ShopNum).Name = Trim(Parse(2))
        Shop(ShopNum).FixesItems = Val(Parse(3))
        
        n = 4
        For i = 1 To MAX_TRADES
            For f = 1 To MAX_GIVE_ITEMS
                Shop(ShopNum).TradeItem(i).GiveItem(f) = Val(Parse((n - 1) + f))
            Next f
            n = n + MAX_GIVE_ITEMS
            For f = 1 To MAX_GIVE_VALUE
                Shop(ShopNum).TradeItem(i).GiveValue(f) = Val(Parse((n - 1) + f))
            Next f
            n = n + MAX_GIVE_VALUE
            For f = 1 To MAX_GET_ITEMS
                Shop(ShopNum).TradeItem(i).GetItem(f) = Val(Parse((n - 1) + f))
            Next f
            n = n + MAX_GET_ITEMS
            For f = 1 To MAX_GET_VALUE
                Shop(ShopNum).TradeItem(i).GetValue(f) = Val(Parse((n - 1) + f))
            Next f
            n = n + MAX_GET_VALUE
            
            Shop(ShopNum).ItemStock(i) = Val(Parse(n))
            n = n + 1
        Next i
        
        ' Save it
        Call SendUpdateShopToAll(ShopNum)
        Call SaveShop(ShopNum)
        Call AddLog(GetPlayerName(Index) & " saving shop #" & ShopNum & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit spell packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditspell" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SPELLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit spell packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "editspell" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(Index, "Invalid Spell Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing spell #" & n & ".", ADMIN_LOG)
        Call SendEditSpellTo(Index, n)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save spell packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "savespell") Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SPELLS Then
            Call HackingAttempt(Index, "Invalid Spell Index")
            Exit Sub
        End If
        
        ' Update the spell
        Spell(n).Name = Parse(2)
        Spell(n).SpellSprite = Val(Parse(3))
        Spell(n).ClassReq = Val(Parse(4))
        Spell(n).LevelReq = Val(Parse(5))
        Spell(n).Type = Val(Parse(6))
        Spell(n).Data1 = Val(Parse(7))
        Spell(n).Data2 = Val(Parse(8))
        Spell(n).Data3 = Val(Parse(9))
                
        ' Save it
        Call SendUpdateSpellToAll(n)
        Call SaveSpell(n)
        Call AddLog(GetPlayerName(Index) & " saving spell #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit skill packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditskill" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "SKILLEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit skill packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "editskill" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SKILLS Then
            Call HackingAttempt(Index, "Invalid Skill Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing skill #" & n & ".", ADMIN_LOG)
        Call SendEditSkillTo(Index, n)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save skill packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "saveskill") Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_SKILLS Then
            Call HackingAttempt(Index, "Invalid Skill Index")
            Exit Sub
        End If
        
        ' Update the skill
        Skill(n).Name = Parse(2)
        Skill(n).SkillSprite = Val(Parse(3))
        Skill(n).ClassReq = Val(Parse(4))
        Skill(n).LevelReq = Val(Parse(5))
        Skill(n).Type = Val(Parse(6))
        Skill(n).Data1 = Val(Parse(7))
        Skill(n).Data2 = Val(Parse(8))
        Skill(n).Data3 = Val(Parse(9))
                
        ' Save it
        Call SendUpdateSkillToAll(n)
        Call SaveSkill(n)
        Call AddLog(GetPlayerName(Index) & " saving Skill #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit quest packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditquest" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "QUESTEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Edit quest packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "editquest" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The spell #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_QUESTS Then
            Call HackingAttempt(Index, "Invalid quest Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing quest #" & n & ".", ADMIN_LOG)
        Call SendEditQuestTo(Index, n)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save quest packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "savequest") Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' Quest #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_QUESTS Then
            Call HackingAttempt(Index, "Invalid Quest Index")
            Exit Sub
        End If
        
        ' Update the quest
        Quest(n).Name = Parse(2)
        Quest(n).SetBy = Val(Parse(3))
        Quest(n).ClassReq = Val(Parse(4))
        Quest(n).LevelMin = Val(Parse(5))
        Quest(n).LevelMax = Val(Parse(6))
        Quest(n).Type = Val(Parse(7))
        Quest(n).Reward = Val(Parse(8))
        Quest(n).RewardValue = Val(Parse(9))
        Quest(n).Data1 = Val(Parse(10))
        Quest(n).Data2 = Val(Parse(11))
        Quest(n).Data3 = Val(Parse(12))
        Quest(n).Description = Parse(13)
                
        ' Save it
        Call SendUpdateQuestToAll(n)
        Call SaveQuest(n)
        Call AddLog(GetPlayerName(Index) & " saving quest #" & n & ".", ADMIN_LOG)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::
    ' :: Accept quest packet ::
    ' :::::::::::::::::::::::::
    If LCase(Parse(0)) = "acceptquest" Then
    Call LogPacket(Trim(Parse(0)))
        'Npc number
        n = Val(Parse(1))
        
        ' Prevent hacking
        If (n < 0) Or (n > MAX_NPCS) Then
            Call HackingAttempt(Index, "Quest Npc Modification")
            Exit Sub
        End If
        
        'Index for the quest
        i = Val(Parse(2))
        
        ' Prevent Hacking
        If (i < 1) Or (i > MAX_NPC_QUESTS) Then
            Call HackingAttempt(Index, "Quest Modification")
            Exit Sub
        End If
        
        ' Check for empty npc quest
        If Npc(n).QuestNPC(i) = 0 Then
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "No Quest Available" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        ' check for duplicate quests
        For f = 1 To MAX_PLAYER_QUESTS
            If Player(Index).Char(Player(Index).CharNum).Quests(f).Num = Npc(n).QuestNPC(i) Then
                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Duplicate Quest" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        Next f
        
        ' Consider quest requirements
        If (Quest(Npc(n).QuestNPC(i)).ClassReq - 1 = GetPlayerClass(Index)) Or (Quest(Npc(n).QuestNPC(i)).ClassReq = 0) Then
            If (GetPlayerLevel(Index) >= Quest(Npc(n).QuestNPC(i)).LevelMin) And (GetPlayerLevel(Index) <= Quest(Npc(n).QuestNPC(i)).LevelMax) Then
                ' check for open quest slot
                f = FindOpenQuestSlot(Index)
                If f > 0 Then
                    Call SetPlayerQuest(Index, f, n, Npc(n).QuestNPC(i))
                    ' Check for quest item already held
                    If Quest(Npc(n).QuestNPC(i)).Type = QUEST_TYPE_FETCH Then
                        If HasItem(Index, Quest(Npc(n).QuestNPC(i)).Data1) > 0 Then
                            Player(Index).Char(Player(Index).CharNum).Quests(f).Count = HasItem(Index, Quest(Npc(n).QuestNPC(i)).Data1)
                            If Player(Index).Char(Player(Index).CharNum).Quests(f).Count > Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2 Then
                                Player(Index).Char(Player(Index).CharNum).Quests(f).Count = Quest(Player(Index).Char(Player(Index).CharNum).Quests(f).Num).Data2
                            End If
                        End If
                    ElseIf Quest(Npc(n).QuestNPC(i)).Type = QUEST_TYPE_KILL Then
                        Player(Index).Char(Player(Index).CharNum).Quests(f).Count = 0
                    End If
                Else
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "No Slots Free" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            Else
                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Level" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        Else
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Wrong Class" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        Call UpdatePlayerQuest(Index, f)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: abandon quest packet ::
    ' ::::::::::::::::::::::::::
    If (LCase(Parse(0)) = "abandonquest") Then
    Call LogPacket(Trim(Parse(0)))
        ' Npc number
        f = Val(Parse(1))
        ' Quest index
        n = Val(Parse(2))
        
        ' Prevent hacking
        If (f < 0) Or (f > MAX_NPCS) Then
            Call HackingAttempt(Index, "Quest Npc Modification")
            Exit Sub
        End If
        
        ' Prevent Hacking
        If n < 0 Or n > MAX_NPC_QUESTS Then
            Call HackingAttempt(Index, "Quest Modification")
        End If
        
        ' Reset quest values
        For i = 1 To MAX_PLAYER_QUESTS
            If Player(Index).Char(Player(Index).CharNum).Quests(i).Num = Npc(f).QuestNPC(n) Then
                Call ClearPlayerQuest(Index, i)
                Call UpdatePlayerQuest(Index, i)
                Exit For
            End If
        Next i
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::
    ' :: Complete quest packet ::
    ' :::::::::::::::::::::::::::
    If LCase(Parse(0)) = "completequest" Then
    Call LogPacket(Trim(Parse(0)))
        ' Npc number
        n = Val(Parse(1))
        
        ' Prevent hacking
        If (n < 0) Or (n > MAX_NPCS) Then
            Call HackingAttempt(Index, "Quest Npc Modification")
            Exit Sub
        End If
        
        'Index for the quest
        i = Val(Parse(2))
        
        ' Prevent Hacking
        If (i < 1) Or (i > MAX_NPC_QUESTS) Then
            Call HackingAttempt(Index, "Quest Modification")
            Exit Sub
        End If
        
        ' Check for completed quest
        For f = 1 To MAX_PLAYER_QUESTS
            If Player(Index).Char(Player(Index).CharNum).Quests(f).Num = Npc(n).QuestNPC(i) Then
                With Player(Index).Char(Player(Index).CharNum).Quests(f)
                    If GetPlayerMap(Index) = .SetMap Then
                        If n = .SetBy Then
                            If .Count = Quest(Npc(n).QuestNPC(i)).Data2 Then
                                ' Take quest items if neccesary
                                If Quest(Npc(n).QuestNPC(i)).Type = QUEST_TYPE_FETCH Then
                                    Call TakeItem(Index, Quest(Npc(n).QuestNPC(i)).Data1, Quest(Npc(n).QuestNPC(i)).Data2)
                                End If
                                ' Give player quest reward
                                If GiveQuestReward(Index, Quest(Npc(n).QuestNPC(i)).Reward, Quest(Npc(n).QuestNPC(i)).RewardValue) = 1 Then
                                    Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Quest Reward" & SEP_CHAR & DarkGrey & SEP_CHAR & END_CHAR)
                                    If Quest(Npc(n).QuestNPC(i)).RewardValue > 1 Then
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Quest(Npc(n).QuestNPC(i)).RewardValue & " x " & Trim(Item(Quest(Npc(n).QuestNPC(i)).Reward).Name) & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                                    Else
                                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Quest(Npc(n).QuestNPC(i)).Reward & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                                    End If
                                    
                                    'Reset player quest
                                    Call ClearPlayerQuest(Index, f)
                                    Call UpdatePlayerQuest(Index, f)
                                End If
                            Else
                                Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Quest Error" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Amount Required" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                            End If
                        Else
                            Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Quest Error" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Quest Npc" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                        End If
                    Else
                        Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Quest Error" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Quest Map" & SEP_CHAR & Grey & SEP_CHAR & END_CHAR)
                    End If
                End With
            End If
        Next f
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::::::::::
    ' :: Request edit GUI packet ::
    ' :::::::::::::::::::::::::::::::
    If LCase(Parse(0)) = "requesteditgui" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        Call SendDataTo(Index, "GUIEDITOR" & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Edit GUI packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "editgui" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' The gui #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_GUIS Then
            Call HackingAttempt(Index, "Invalid GUI Index")
            Exit Sub
        End If
        
        Call AddLog(GetPlayerName(Index) & " editing GUI #" & n & ".", ADMIN_LOG)
        Call TextAdd(frmCServer.txtText, GetPlayerName(Index) & " editing GUI #" & n & ".", True)
        Call SendEditGUITo(Index, n)
    End If
    
    ' :::::::::::::::::::::::
    ' :: Save GUI packet ::
    ' :::::::::::::::::::::::
    If (LCase(Parse(0)) = "savegui") Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_DEVELOPER Then
            Call HackingAttempt(Index, "Admin Cloning")
            Exit Sub
        End If
        
        ' GUI #
        n = Val(Parse(1))
        
        ' Prevent hacking
        If n < 0 Or n > MAX_GUIS Then
            Call HackingAttempt(Index, "Invalid GUI Index")
            Exit Sub
        End If
        
        ' Update the GUI
        GUI(n).Name = Parse(2)
        GUI(n).Designer = Parse(3)
        GUI(n).Revision = Val(Parse(4))
        
        f = 5
        
        For i = 1 To 7
            GUI(n).Background(i).Data1 = Val(Parse(f))
            GUI(n).Background(i).Data2 = Val(Parse(f + 1))
            GUI(n).Background(i).Data3 = Val(Parse(f + 2))
            GUI(n).Background(i).Data4 = Val(Parse(f + 3))
            GUI(n).Background(i).Data5 = Val(Parse(f + 4))
            
            f = f + 5
        Next i
        For i = 1 To 5
            GUI(n).Menu(i).Data1 = Val(Parse(f))
            GUI(n).Menu(i).Data2 = Val(Parse(f + 1))
            GUI(n).Menu(i).Data3 = Val(Parse(f + 2))
            GUI(n).Menu(i).Data4 = Val(Parse(f + 3))
            
            f = f + 4
        Next i
        For i = 1 To 4
            GUI(n).Login(i).Data1 = Val(Parse(f))
            GUI(n).Login(i).Data2 = Val(Parse(f + 1))
            GUI(n).Login(i).Data3 = Val(Parse(f + 2))
            GUI(n).Login(i).Data4 = Val(Parse(f + 3))
            
            f = f + 4
        Next i
        For i = 1 To 4
            GUI(n).NewAcc(i).Data1 = Val(Parse(f))
            GUI(n).NewAcc(i).Data2 = Val(Parse(f + 1))
            GUI(n).NewAcc(i).Data3 = Val(Parse(f + 2))
            GUI(n).NewAcc(i).Data4 = Val(Parse(f + 3))
            
            f = f + 4
        Next i
        For i = 1 To 4
            GUI(n).DelAcc(i).Data1 = Val(Parse(f))
            GUI(n).DelAcc(i).Data2 = Val(Parse(f + 1))
            GUI(n).DelAcc(i).Data3 = Val(Parse(f + 2))
            GUI(n).DelAcc(i).Data4 = Val(Parse(f + 3))
            
            f = f + 4
        Next i
        For i = 1 To 2
            GUI(n).Credits(i).Data1 = Val(Parse(f))
            GUI(n).Credits(i).Data2 = Val(Parse(f + 1))
            GUI(n).Credits(i).Data3 = Val(Parse(f + 2))
            GUI(n).Credits(i).Data4 = Val(Parse(f + 3))
            
            f = f + 4
        Next i
        For i = 1 To 5
            GUI(n).Chars(i).Data1 = Val(Parse(f))
            GUI(n).Chars(i).Data2 = Val(Parse(f + 1))
            GUI(n).Chars(i).Data3 = Val(Parse(f + 2))
            GUI(n).Chars(i).Data4 = Val(Parse(f + 3))
            
            f = f + 4
        Next i
        For i = 1 To 14
            GUI(n).NewChar(i).Data1 = Val(Parse(f))
            GUI(n).NewChar(i).Data2 = Val(Parse(f + 1))
            GUI(n).NewChar(i).Data3 = Val(Parse(f + 2))
            GUI(n).NewChar(i).Data4 = Val(Parse(f + 3))
            
            f = f + 4
        Next i
        
        ' Save it
        Call SendUpdateGUIToAll(n)
        Call SaveGUI(n)
        Call AddLog(GetPlayerName(Index) & " saving GUI #" & n & ".", ADMIN_LOG)
        Call TextAdd(frmCServer.txtText, GetPlayerName(Index) & " saving GUI #" & n & ".", True)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Set access packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "setaccess" Then
    Call LogPacket(Trim(Parse(0)))
        ' Prevent hacking
        If GetPlayerAccess(Index) < ADMIN_CREATOR Then
            Call HackingAttempt(Index, "Trying to use powers not available")
            Exit Sub
        End If
        
        ' The index
        n = FindPlayer(Parse(1))
        ' The access
        i = Val(Parse(2))
        
        
        ' Check for invalid access level
        If i >= 3 Or i <= 5 Then
            ' Check if player is on
            If n > 0 Then
                If GetPlayerAccess(n) <= 0 Then
                    Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
                End If
                
                Call SetPlayerAccess(n, i)
                Call SendPlayerData(n)
                Call AddLog(GetPlayerName(Index) & " has modified " & GetPlayerName(n) & "'s access.", ADMIN_LOG)
            Else
                Call PlayerMsg(Index, "Player is not online.", White)
            End If
        Else
            Call PlayerMsg(Index, "Invalid access level.", Red)
        End If
                
        Exit Sub
    End If
    
    ' :::::::::::::::::::::::
    ' :: Who online packet ::
    ' :::::::::::::::::::::::
    If LCase(Parse(0)) = "whosonline" Then
    Call LogPacket(Trim(Parse(0)))
        Call SendWhosOnline(Index)
        Exit Sub
    End If
    
    ' ::::::::::::::::::::::::::
    ' :: Trade request packet ::
    ' ::::::::::::::::::::::::::
    If LCase(Parse(0)) = "traderequest" Then
    Call LogPacket(Trim(Parse(0)))
        ' Trade num
        n = Val(Parse(2))
        
        ' Prevent hacking
        If (n <= 0) Or (n > MAX_TRADES) Then
            Call HackingAttempt(Index, "Trade Request Modification")
            Exit Sub
        End If
        
        ' Index for shop
        i = Val(Parse(1))
        
        ' Check if inv full
        For f = 1 To MAX_GET_ITEMS
            If Shop(i).TradeItem(n).GetItem(f) > 0 Then
                x = FindOpenInvSlot(Index, Shop(i).TradeItem(n).GetItem(f))
                If x = 0 Then
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Inventory Full" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
        Next f
        
        ' Check if they have the items
        For f = 1 To MAX_GIVE_ITEMS
            If Shop(i).TradeItem(n).GiveItem(f) > 0 Then
                If HasItem(Index, Shop(i).TradeItem(n).GiveItem(f)) >= Shop(i).TradeItem(n).GiveValue(f) Then
                Else
                    Call SendDataTo(Index, "BLITWARNMSG" & SEP_CHAR & "Need More" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                    Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Trim(Item(Shop(i).TradeItem(n).GiveItem(f)).Name) & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
        Next f
        
        ' If they have all the items, make the trade
        For f = 1 To MAX_GIVE_ITEMS
            Call TakeItem(Index, Shop(i).TradeItem(n).GiveItem(f), Shop(i).TradeItem(n).GiveValue(f))
        Next f
        For f = 1 To MAX_GET_ITEMS
            Call GiveItem(Index, Shop(i).TradeItem(n).GetItem(f), Shop(i).TradeItem(n).GetValue(f))
        Next f
        Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Trade Successful" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
        Exit Sub
    End If
    
    ' :::::::::::::::::::::
    ' :: Fix item packet ::
    ' :::::::::::::::::::::
    If LCase(Parse(0)) = "fixitem" Then
    Call LogPacket(Trim(Parse(0)))
        ' Inv num
        n = Val(Parse(1))
        
        ' Make sure its a equipable item
        If Item(GetPlayerInvItemNum(Index, n)).Type < ITEM_TYPE_WEAPON Or Item(GetPlayerInvItemNum(Index, n)).Type > ITEM_TYPE_SHIELD Then
            If Item(GetPlayerInvItemNum(Index, n)).Type <> ITEM_TYPE_TOOL Then
                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Item Not Fixable" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        End If
        
        ' Check if they have a full inventory
        If FindOpenInvSlot(Index, GetPlayerInvItemNum(Index, n)) <= 0 Then
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "No Space Left" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        ' Now check the rate of pay
        ItemNum = GetPlayerInvItemNum(Index, n)
        i = Int(Item(GetPlayerInvItemNum(Index, n)).Data2 / 5)
        If i <= 0 Then i = 1
        
        DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(Index, n)
        GoldNeeded = Int(DurNeeded * i / 2)
        If GoldNeeded <= 0 Then GoldNeeded = 1
        
        ' Check if they even need it repaired
        If DurNeeded <= 0 Then
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Item Okay" & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
            Exit Sub
        End If
        
        ' Check if they have enough for at least one point
        If HasItem(Index, 1) >= i Then
            ' Check if they have enough for a total restoration
            If HasItem(Index, 1) >= GoldNeeded Then
                Call TakeItem(Index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(Index, n, Item(ItemNum).Data1)
                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Trim(Item(ItemNum).Name) & ": +" & DurNeeded & SEP_CHAR & BrightGreen & SEP_CHAR & END_CHAR)
                Call SendInventoryUpdate(Index, n)
            Else
                ' They dont so restore as much as we can
                DurNeeded = (HasItem(Index, 1) / i)
                GoldNeeded = Int(DurNeeded * i / 2)
                If GoldNeeded <= 0 Then GoldNeeded = 1
                
                Call TakeItem(Index, 1, GoldNeeded)
                Call SetPlayerInvItemDur(Index, n, GetPlayerInvItemDur(Index, n) + DurNeeded)
                Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & Item(ItemNum).Name & ": +" & DurNeeded & SEP_CHAR & BrightGreen & SEP_CHAR & END_CHAR)
                Call SendInventoryUpdate(Index, n)
            End If
        Else
            Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Insufficient Currency" & SEP_CHAR & BrightRed & SEP_CHAR & END_CHAR)
        End If
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Search packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "search" Then
    Call LogPacket(Trim(Parse(0)))
        x = Val(Parse(1))
        y = Val(Parse(2))
        
        ' Prevent subscript out of range
        If x < 0 Or x > MAX_MAPX Or y < 0 Or y > MAX_MAPY Then
            Exit Sub
        End If
        
        ' Check for a player
        For i = 1 To HighIndex
            If IsPlaying(i) And GetPlayerMap(Index) = GetPlayerMap(i) And GetPlayerX(i) = x And GetPlayerY(i) = y Then
                
                '' Consider the player
                'If GetPlayerLevel(i) >= GetPlayerLevel(Index) + 5 Then
                    'Call PlayerMsg(Index, "You wouldn't stand a chance.", BrightRed)
                'Else
                    'If GetPlayerLevel(i) > GetPlayerLevel(Index) Then
                        'Call PlayerMsg(Index, "This one seems to have an advantage over you.", Yellow)
                    'Else
                        'If GetPlayerLevel(i) = GetPlayerLevel(Index) Then
                            'Call PlayerMsg(Index, "This would be an even fight.", White)
                        'Else
                            'If GetPlayerLevel(Index) >= GetPlayerLevel(i) + 5 Then
                                'Call PlayerMsg(Index, "You could slaughter that player.", BrightBlue)
                            'Else
                                'If GetPlayerLevel(Index) > GetPlayerLevel(i) Then
                                    'Call PlayerMsg(Index, "You would have an advantage over that player.", Yellow)
                                'End If
                            'End If
                        'End If
                    'End If
                'End If
            
                ' Change target
                Player(Index).Target = i
                Player(Index).TargetType = TARGET_TYPE_PLAYER
                'Call PlayerMsg(Index, "Your target is now " & GetPlayerName(i) & ".", Yellow)
                'Call SendDataTo(Index, "BLITPLAYERMSG" & SEP_CHAR & "Target: " & GetPlayerName(i) & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                Call SendDataTo(Index, "BLITPKMSG" & SEP_CHAR & "Target: " & GetPlayerName(i) & SEP_CHAR & i & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                Exit Sub
            End If
        Next i
        
        ' Check for a resource
        For i = 1 To MAX_MAP_RESOURCES
            If MapResource(GetPlayerMap(Index), i).Num > 0 Then
                If MapResource(GetPlayerMap(Index), i).x = x And MapResource(GetPlayerMap(Index), i).y = y Then
                    ' Change target
                    Player(Index).Target = i
                    Player(Index).TargetType = TARGET_TYPE_RESOURCE
                    Call SendDataTo(Index, "BLITRESOURCEMSG" & SEP_CHAR & i & SEP_CHAR & "Target: " & Trim(Npc(MapResource(GetPlayerMap(Index), i).Num).Name) & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
        Next i
        
        ' Check for an npc
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(GetPlayerMap(Index), i).Num > 0 Then
                If MapNpc(GetPlayerMap(Index), i).x = x And MapNpc(GetPlayerMap(Index), i).y = y Then
                    ' Change target
                    Player(Index).Target = i
                    Player(Index).TargetType = TARGET_TYPE_NPC
                    Call SendDataTo(Index, "BLITNPCMSG" & SEP_CHAR & "Target: " & Trim(Npc(MapNpc(GetPlayerMap(Index), i).Num).Name) & SEP_CHAR & i & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
        Next i
        
        ' Check for an item
        For i = MAX_MAP_ITEMS To 1 Step -1
            If MapItem(GetPlayerMap(Index), i).Num > 0 Then
                If MapItem(GetPlayerMap(Index), i).x = x And MapItem(GetPlayerMap(Index), i).y = y Then
                    ' Change target
                    Player(Index).Target = i
                    Player(Index).TargetType = TARGET_TYPE_ITEM
                    Call SendDataTo(Index, "BLITITEMMSG" & SEP_CHAR & "Target: " & Trim(Item(MapItem(GetPlayerMap(Index), i).Num).Name) & SEP_CHAR & i & SEP_CHAR & Yellow & SEP_CHAR & END_CHAR)
                    Exit Sub
                End If
            End If
        Next i
        
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Spells packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "spells" Then
    Call LogPacket(Trim(Parse(0)))
        Call SendPlayerSpells(Index)
        Exit Sub
    End If
    
    ' :::::::::::::::::::
    ' :: Skills packet ::
    ' :::::::::::::::::::
    If LCase(Parse(0)) = "skills" Then
    Call LogPacket(Trim(Parse(0)))
        Call SendPlayerSkills(Index)
        Exit Sub
    End If
    
    ' :::::::::::::::::
    ' :: Cast packet ::
    ' :::::::::::::::::
    If LCase(Parse(0)) = "cast" Then
    Call LogPacket(Trim(Parse(0)))
        ' Spell slot
        n = Val(Parse(1))
        
        Call CastSpell(Index, n)
        
        Exit Sub
    End If
    
End Sub
