Attribute VB_Name = "modInput"
Option Explicit

' ******************************************
' **               rootSource               **
' ******************************************

' keyboard input declares
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetForegroundWindow Lib "user32" () As Long

' player input text buffer
Public MyText As String

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Key constants
Private Const VK_UP As Long = &H26
Private Const VK_DOWN As Long = &H28
Private Const VK_LEFT As Long = &H25
Private Const VK_RIGHT As Long = &H27
Private Const VK_SHIFT As Long = &H10
'Private Const VK_RETURN As Long = &HD ' not used
Private Const VK_CONTROL As Long = &H11

Public Sub CheckInputKeys()

    ' Check to make sure they aren't trying to auto do anything
    If GetAsyncKeyState(VK_UP) >= 0 Then DirUp = False
    If GetAsyncKeyState(VK_DOWN) >= 0 Then DirDown = False
    If GetAsyncKeyState(VK_LEFT) >= 0 Then DirLeft = False
    If GetAsyncKeyState(VK_RIGHT) >= 0 Then DirRight = False
    If GetAsyncKeyState(VK_CONTROL) >= 0 Then ControlDown = False
    If GetAsyncKeyState(VK_SHIFT) >= 0 Then ShiftDown = False

    If frmMainGame.WindowState = vbMinimized Then Exit Sub

    If GetKeyState(vbKeyShift) < 0 Then
        ShiftDown = True
    Else
        ShiftDown = False
    End If
    
    If GetKeyState(vbKeyReturn) < 0 Then
        CheckMapGetItem
    End If
    
    If GetKeyState(vbKeyControl) < 0 Then
        ControlDown = True
    Else
        ControlDown = False
    End If
    
    'Move Up
    If GetKeyState(vbKeyUp) < 0 Then
        DirUp = True
        DirDown = False
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirUp = False
    End If
    
    'Move Right
    If GetKeyState(vbKeyRight) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = True
        Exit Sub
    Else
        DirRight = False
    End If
    
    'Move down
    If GetKeyState(vbKeyDown) < 0 Then
        DirUp = False
        DirDown = True
        DirLeft = False
        DirRight = False
        Exit Sub
    Else
        DirDown = False
    End If
    
    'Move left
    If GetKeyState(vbKeyLeft) < 0 Then
        DirUp = False
        DirDown = False
        DirLeft = True
        DirRight = False
        Exit Sub
    Else
        DirLeft = False
    End If
        
End Sub

Private Sub CheckMapGetItem()
Dim Buffer As clsBuffer
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 Then
        If Trim$(MyText) = vbNullString Then
            Set Buffer = New clsBuffer
            Buffer.PreAllocate 2
            Buffer.WriteInteger CMapGetItem
            Player(MyIndex).MapGetTimer = GetTickCount
            Call SendData(Buffer.ToArray())
        End If
    End If
End Sub

' Processes input from player
Public Sub HandleKeyPresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim Name As String
Dim i As Long
Dim n As Long
Dim Command() As String
Dim Buffer As clsBuffer

    ChatText = Trim$(MyText)
    
    If LenB(ChatText) = 0 Then Exit Sub
    
    MyText = LCase$(ChatText)
    
    ' Handle when the player presses the return key
    If KeyAscii = vbKeyReturn Then
    
        ' Broadcast message
        If Left$(ChatText, 1) = "'" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
            If Len(ChatText) > 0 Then
                Call BroadcastMsg(ChatText)
            End If
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        End If
        
        ' Emote message
        If Left$(ChatText, 1) = "-" Then
            MyText = Mid$(ChatText, 2, Len(ChatText) - 1)
            If Len(ChatText) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        End If
        
        ' Player message
        If Left$(ChatText, 1) = "!" Then
            ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
            Name = vbNullString
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid$(ChatText, i, 1) <> Space(1) Then
                    Name = Name & Mid$(ChatText, i, 1)
                Else
                    Exit For
                End If
            Next
            
            ChatText = Mid$(ChatText, i, Len(ChatText) - 1)
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                MyText = Mid$(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, Name)
            Else
                Call AddText("Usage: !playername (message)", AlertColor)
            End If
            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        End If
        
        ' Global Message
        If Left$(ChatText, 1) = vbQuote Then
            If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString
                Exit Sub
            End If
        End If
    
        ' Admin Message
        If Left$(ChatText, 1) = "=" Then
            If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then
                ChatText = Mid$(ChatText, 2, Len(ChatText) - 1)
                If Len(ChatText) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                MyText = vbNullString
                frmMainGame.txtMyChat.Text = vbNullString
                Exit Sub
            End If
        End If
        
        If Left$(MyText, 1) = "/" Then
            Command = Split(MyText, Space(1))
            
            Select Case Command(0)
                Case "/help"
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("'msghere = Broadcast Message", HelpColor)
                    Call AddText("-msghere = Emote Message", HelpColor)
                    Call AddText("!namehere msghere = Player Message", HelpColor)
                    Call AddText("Available Commands: /help, /info, /who, /fps, /inv, /stats, /train, /trade, /party, /join, /leave, /resetui", HelpColor)
                                    
                Case "/info"
                    ' Checks to make sure we have more than one string in the array
                    If UBound(Command) < 1 Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /info (name)", AlertColor
                        GoTo Continue
                    End If
                    Set Buffer = New clsBuffer
                    Buffer.PreAllocate Len(Command(1)) + 4
                    Buffer.WriteInteger CPlayerInfoRequest
                    Buffer.WriteString Command(1)
                    Call SendData(Buffer.ToArray())
            
                ' Whos Online
                Case "/who"
                    SendWhosOnline
                                
                ' Checking fps
                Case "/fps"
                    BFPS = Not BFPS
                    
                Case "/guildinvite"
                    If UBound(Command) < 1 Then
                        AddText "Usage: /guildinvite (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendGuildInvite Command(1)
                    
                Case "/guildkick"
                    If UBound(Command) < 1 Then
                        AddText "Usage: /guildkick (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendGuildKick Command(1)
                                            
                Case "/guildpromote"
                    If UBound(Command) < 2 Then
                        AddText "Usage: /guildkick (name) (access)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Val(Command(2))) = False Then
                        AddText "Usage: /guildkick (name) (access)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendGuildPromote Command(1), Val(Command(2))
                ' Show inventory
                Case "/inv"
                    UpdateInventory
                    frmMainGame.picInvList.Visible = (Not frmMainGame.picInvList.Visible)
                
                ' Request stats
                Case "/stats"
                    Set Buffer = New clsBuffer
                    Buffer.PreAllocate 2
                    Buffer.WriteInteger CGetStats
                    SendData Buffer.ToArray()
            
                ' Show training
                Case "/train"
                    frmTraining.Show vbModal
                            
                ' Request stats
                Case "/trade"
                    Set Buffer = New clsBuffer
                    Buffer.PreAllocate 2
                    Buffer.WriteInteger CTrade
                    SendData Buffer.ToArray()
                    
                ' Party request
                Case "/party"
                    ' Make sure they are actually sending something
                    If UBound(Command) < 1 Then
                        AddText "Usage: /party (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /party (name)", AlertColor
                        GoTo Continue
                    End If
                        
                    Call SendPartyRequest(Command(1))
                    
                ' Join party
                Case "/join"
                    SendJoinParty
                
                ' Leave party
                Case "/leave"
                    SendLeaveParty
                    
                Case "/resetui"
                    Call ResetUI
                
                ' // Monitor Admin Commands //
                ' Admin Help
                Case "/admin"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    Call AddText("Social Commands:", HelpColor)
                    Call AddText("""msghere = Global Admin Message", HelpColor)
                    Call AddText("=msghere = Private Admin Message", HelpColor)
                    Call AddText("Available Commands: /admin, /loc, /mapeditor, /warpmeto, /warptome, /warpto, /setsprite, /mapreport, /kick, /ban, /edititem, /respawn, /editnpc, /motd, /editshop, /editspell, /debug", HelpColor)
                
                ' Kicking a player
                Case "/kick"
                    If GetPlayerAccess(MyIndex) < ADMIN_MONITOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /kick (name)", AlertColor
                        GoTo Continue
                    End If

                    SendKick Command(1)
                            
                ' // Mapper Admin Commands //
                ' Location
                Case "/loc"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    BLoc = Not BLoc
                
                ' Map Editor
                Case "/mapeditor"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditMap
                
                ' Warping to a player
                Case "/warpmeto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warpmeto (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    WarpMeTo Command(1)
                            
                ' Warping a player to you
                Case "/warptome"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Then
                        AddText "Usage: /warptome (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    WarpToMe Command(1)
                            
                ' Warping to a map
                Case "/warpto"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /warpto (map #)", AlertColor
                        GoTo Continue
                    End If
                    
                    n = CLng(Command(1))
                
                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If
                
                ' Setting sprite
                Case "/setsprite"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If
                    
                    If Not IsNumeric(Command(1)) Then
                        AddText "Usage: /setsprite (sprite #)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendSetSprite CLng(Command(1))
                
                ' Map report
                Case "/mapreport"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    Set Buffer = New clsBuffer
                    Buffer.PreAllocate 2
                    Buffer.WriteInteger CMapReport
                    SendData Buffer.ToArray()
            
                ' Respawn request
                Case "/respawn"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendMapRespawn
            
                ' MOTD change
                Case "/motd"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /motd (new motd)", AlertColor
                        GoTo Continue
                    End If

                    SendMOTDChange Right$(ChatText, Len(ChatText) - 5)
                
                ' Check the ban list
                Case "/banlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendBanList
                
                ' Banning a player
                Case "/ban"
                    If GetPlayerAccess(MyIndex) < ADMIN_MAPPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    If UBound(Command) < 1 Then
                        AddText "Usage: /ban (name)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendBan Command(1)
                
                ' // Developer Admin Commands //
                ' Editing item request
                Case "/edititem"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If

                    SendRequestEditItem
                
                ' Editing npc request
                Case "/editnpc"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditNpc
                
                ' Editing shop request
                Case "/editshop"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditShop
            
                ' Editing spell request
                Case "/editspell"
                    If GetPlayerAccess(MyIndex) < ADMIN_DEVELOPER Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRequestEditSpell
                
                ' // Creator Admin Commands //
                Case "/createguild"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    
                    If UBound(Command) < 2 Then
                        AddText "Usage: /createguild (user) (guild)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendCreateGuild Command(1), Command(2)
                    
                Case "/removefromguild"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    
                    If UBound(Command) < 2 Then
                        AddText "Usage: /removefromguild (user) (guild)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendRemoveFromGuild Command(1), Command(2)
                
                ' Giving another player access
                Case "/setaccess"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    
                    If UBound(Command) < 2 Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If
                    
                    If IsNumeric(Command(1)) Or Not IsNumeric(Command(2)) Then
                        AddText "Usage: /setaccess (name) (access)", AlertColor
                        GoTo Continue
                    End If
                    
                    SendSetAccess Command(1), CLng(Command(2))
                
                ' Ban destroy
                Case "/destroybanlist"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    SendBanDestroy
                    
                ' Packet debug mode
                Case "/debug"
                    If GetPlayerAccess(MyIndex) < ADMIN_CREATOR Then
                        AddText "You need to be a high enough staff member to do this!", AlertColor
                        GoTo Continue
                    End If
                    
                    DebugMode = (Not DebugMode)
                
                Case Else
                    AddText "Not a valid command!", HelpColor
                    
            End Select
            
'continue label where we go instead of exiting the sub
Continue:

            MyText = vbNullString
            frmMainGame.txtMyChat.Text = vbNullString
            Exit Sub
        End If
        
        ' And if neither, then add the character to the user's text buffer
        'If (KeyAscii <> vbKeyReturn) Then
        '    If (KeyAscii <> vbKeyBack) Then
        '
        '        ' Make sure the character is on standard English keyboard
        '        If KeyAscii >= 32 Then ' Asc(" ")
        '            If KeyAscii <= 126 Then ' Asc("~")
        '                MyText = MyText & ChrW$(KeyAscii)
        '            End If
        '        End If
        '
        '    End If
        'End If
    
        ' Handle when the user presses the backspace key
        'If (KeyAscii = vbKeyBack) Then
        '    If LenB(MyText) > 0 Then MyText = Mid$(MyText, 1, Len(MyText) - 1)
        'End If
        
        ' Say message
        If Len(ChatText) > 0 Then
            Call SayMsg(ChatText)
        End If
        
        MyText = vbNullString
        frmMainGame.txtMyChat.Text = vbNullString
        Exit Sub
    End If
    
    
End Sub

