Attribute VB_Name = "modText"
Option Explicit

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Public Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CelWidth As Long
    CelHeight As Long
    BaseCharOffset As Byte
    CharWidth(0 To 255) As Byte
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH
    Texture As DX8TextureRec
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
End Type

Public Font_Default As CustomFont
Public Font_Georgia As CustomFont

Public Const FVF_SIZE As Long = 28

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal Alpha As Long = 255, Optional Shadow As Boolean = True)
    Dim TempVA(0 To 3)  As TLVERTEX
    Dim TempVAS(0 To 3) As TLVERTEX
    Dim TempStr() As String
    Dim Count As Integer
    Dim Ascii() As Byte
    Dim Row As Integer
    Dim u As Single
    Dim v As Single
    Dim I As Long
    Dim J As Long
    Dim KeyPhrase As Byte
    Dim TempColor As Long
    Dim ResetColor As Byte
    Dim srcRect As RECT
    Dim v2 As D3DVECTOR2
    Dim v3 As D3DVECTOR2
    Dim yOffset As Single

    ' Set the color
    Color = DX8Color(Color, Alpha)
    
    ' Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    
    ' Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    
    ' Set the temp color (or else the first character has no color)
    TempColor = Color
    
    If gTexture(UseFont.Texture.Texture).Loaded = False Then
        SetTexture UseFont.Texture
    Else
        gTexture(UseFont.Texture.Texture).Timer = timeGetTime + 150000
    End If
    
    ' Set the texture
    Direct3D_Device.SetTexture 0, gTexture(UseFont.Texture.Texture).Texture
    
    ' Loop through each line if there are line breaks (vbCrLf)
    Dim Size As Long
    Size = UBound(TempStr)
    For I = 0 To Size
        If Len(TempStr(I)) > 0 Then
            yOffset = I * UseFont.CharHeight
            Count = 0
            ' Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(I), vbFromUnicode)
            
            Dim size2 As Long
            size2 = Len(TempStr(I))
            ' Loop through the characters
            For J = 1 To size2
                ' Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(J - 1)).Vertex(0), FVF_SIZE * 4)
                
                ' Set up the verticies
                TempVA(0).X = X + Count
                TempVA(0).Y = Y + yOffset
                TempVA(1).X = TempVA(1).X + X + Count
                TempVA(1).Y = TempVA(0).Y
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                
                ' Set the colors
                TempVA(0).Color = TempColor
                TempVA(1).Color = TempColor
                TempVA(2).Color = TempColor
                TempVA(3).Color = TempColor
                
                ' Draw the verticies
                Call Direct3D_Device.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0)))
                
                ' Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(Ascii(J - 1))
                
                ' Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
            Next J
        End If
    Next I
End Sub

Sub EngineInitFontTextures()
    ' Default
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Default.Texture.Texture = NumTextures
    Font_Default.Texture.filepath = App.Path & FONT_PATH & "texdefault.png"
    LoadTexture Font_Default.Texture
    
    ' Georgia
    NumTextures = NumTextures + 1
    ReDim Preserve gTexture(NumTextures)
    Font_Georgia.Texture.Texture = NumTextures
    Font_Georgia.Texture.filepath = App.Path & FONT_PATH & "georgia.png"
    LoadTexture Font_Georgia.Texture
End Sub

Sub UnloadFontTextures()
    UnloadFont Font_Default
    UnloadFont Font_Georgia
End Sub

Sub UnloadFont(font As CustomFont)
    font.Texture.Texture = 0
End Sub

Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal FileName As String)
    Dim FileNum As Byte
    Dim LoopChar As Long
    Dim Row As Single
    Dim u As Single
    Dim v As Single

    ' Load the header information
    FileNum = FreeFile
    Open App.Path & FONT_PATH & FileName For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    ' Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CelHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CelWidth
    theFont.ColFactor = theFont.HeaderInfo.CelWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CelHeight / theFont.HeaderInfo.BitmapHeight
    
    ' Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        ' tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        ' Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0) ' Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).TU = u
            .Vertex(0).TV = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).TU = u + theFont.ColFactor
            .Vertex(1).TV = v
            .Vertex(1).X = theFont.HeaderInfo.CelWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).TU = u
            .Vertex(2).TV = v + theFont.RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = theFont.HeaderInfo.CelHeight
            .Vertex(2).Z = 0
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).TU = u + theFont.ColFactor
            .Vertex(3).TV = v + theFont.RowFactor
            .Vertex(3).X = theFont.HeaderInfo.CelWidth
            .Vertex(3).Y = theFont.HeaderInfo.CelHeight
            .Vertex(3).Z = 0
        End With
    Next LoopChar
End Sub

Sub EngineInitFontSettings()
    LoadFontHeader Font_Default, "texdefault.dat"
    LoadFontHeader Font_Georgia, "georgia.dat"
End Sub

Public Function DX8Color(ByVal ColorNum As Long, Optional ByVal Alpha As Long = 255) As Long
    Select Case ColorNum
        Case 0 ' Black
            DX8Color = D3DColorARGB(Alpha, 0, 0, 0)
        Case 1 ' Blue
            DX8Color = D3DColorARGB(Alpha, 16, 104, 237)
        Case 2 ' Green
            DX8Color = D3DColorARGB(Alpha, 119, 188, 84)
        Case 3 ' Cyan
            DX8Color = D3DColorARGB(Alpha, 16, 224, 237)
        Case 4 ' Red
            DX8Color = D3DColorARGB(Alpha, 201, 0, 0)
        Case 5 ' Magenta
            DX8Color = D3DColorARGB(Alpha, 255, 0, 255)
        Case 6 ' Brown
            DX8Color = D3DColorARGB(Alpha, 175, 149, 92)
        Case 7 ' Grey
            DX8Color = D3DColorARGB(Alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            DX8Color = D3DColorARGB(Alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            DX8Color = D3DColorARGB(Alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            DX8Color = D3DColorARGB(Alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            DX8Color = D3DColorARGB(Alpha, 157, 242, 242)
        Case 12 ' BrightRed
            DX8Color = D3DColorARGB(Alpha, 255, 0, 0)
        Case 13 ' Pink
            DX8Color = D3DColorARGB(Alpha, 255, 118, 221)
        Case 14 ' Yellow
            DX8Color = D3DColorARGB(Alpha, 255, 255, 0)
        Case 15 ' White
            DX8Color = D3DColorARGB(Alpha, 255, 255, 255)
        Case 16 ' Dark brown
            DX8Color = D3DColorARGB(Alpha, 98, 84, 52)
        Case 17 ' Orange
            DX8Color = D3DColorARGB(Alpha, 255, 96, 0)
        Case Else
            DX8Color = ColorNum
    End Select
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Integer
    Dim LoopI As Integer

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    Dim Size As Long
    Size = Len(text)
    For LoopI = 1 To Size
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
    Next LoopI
End Function

Public Sub DrawPlayerName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim Color As Long
    Dim Difference As Long
    Dim text, Guild, Level As String

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = White
            Case 1
                Color = Cyan
            Case 2
                Color = Green
            Case 3
                Color = Blue
            Case 4
                Color = Yellow
            Case 5
                Color = Orange
        End Select
    Else
        If GetPlayerPK(Index) = PLAYER_KILLER Then
            Color = BrightRed
        Else
            Color = BrightBlue
        End If
    End If

    text = Trim$(Player(Index).Name)
    TextX = GetPlayerTextX(Index) - (EngineGetTextWidth(Font_Default, text) / 2)
    
    If Options.Levels = 1 Then
        'If Not Index = MyIndex Then
            Level = GetPlayerLevel(Index)
            TextX = TextX + (EngineGetTextWidth(Font_Default, Level) / 2) + 4
        'End If
    End If
    
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = GetPlayerTextY(Index)
    Else
        If (Tex_Character(GetPlayerSprite(Index)).Height / 4) < 32 Then
            TextY = TextY - 16
        End If
    
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + TempPlayer(Index).yOffset - (Tex_Character(GetPlayerSprite(Index)).Height / 4) + 16
        Guild = Trim$(Player(Index).Guild)
        
        If Options.Guilds = 1 Then
            If Len(Guild) > 0 Then
                TextY = TextY - 8
            End If
        End If
    End If

    ' Draw name
    RenderText Font_Default, text, TextX, TextY, Color
    
    If Options.Levels = 1 Then
        'If Not Index = MyIndex Then
            TextX = TextX - (EngineGetTextWidth(Font_Default, (Trim$(Level)))) - 4

            If GetPlayerLevel(Index) = GetPlayerLevel(MyIndex) Then
                Color = D3DColorARGB(255, 255, 255, 0)
            ElseIf GetPlayerLevel(Index) > GetPlayerLevel(MyIndex) Then
                Difference = GetPlayerLevel(Index) - GetPlayerLevel(MyIndex)
                If Difference > 10 Then Difference = 10
                Color = D3DColorARGB(255, 255, 255 - (255 * (Difference / 10)), 0)
            ElseIf GetPlayerLevel(Index) < GetPlayerLevel(MyIndex) Then
                Difference = GetPlayerLevel(MyIndex) - GetPlayerLevel(Index)
                If Difference > 10 Then Difference = 10
                Color = D3DColorARGB(255, 255 - (255 * (Difference / 10)), 255, 0)
            End If
            
            ' Draw Level
            RenderText Font_Default, Level, TextX, TextY, Color
        'End If
    End If
    
    If Options.Titles = 1 And Player(Index).CurTitle > 0 Then
        text = Trim$(title(Player(Index).CurTitle).Name)
        Color = Trim$(title(Player(Index).CurTitle).Color)
        
        TextX = GetPlayerTextX(Index) - EngineGetTextWidth(Font_Default, Trim$(title(Player(Index).CurTitle).Name)) / 2
        TextY = TextY - 12
        
        ' Draw title
        RenderText Font_Default, text, TextX, TextY, Color
    End If
    
    If Options.Guilds = 1 Then
        If Len(Guild) > 0 Then
            Guild = "<" & Guild & ">"
    
            Select Case Player(Index).GuildAcc
                Case 1
                    Color = White
                Case 2
                    Color = BrightGreen
                Case 3
                    Color = BrightBlue
                Case 4
                    Color = Yellow
            End Select
    
            ' Re-center
            TextX = GetPlayerTextX(Index) + (EngineGetTextWidth(Font_Default, Guild) / 2)
            
            ' Determine location for text
            If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
                TextY = GetPlayerTextY(Index)
            Else
                TextY = GetPlayerTextY(Index) - (Tex_Character(GetPlayerSprite(Index)).Height / 4)
            End If
            
            If Options.Titles = 0 Or Player(Index).CurTitle = 0 Then
                TextY = TextY + 8
            End If
            
            ' Draw Guild
            RenderText Font_Default, Guild, TextX, TextY, Color
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawNPCName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim Color As Long
    Dim npcNum As Long
    Dim Name As String
    Dim Level As String
    Dim Difference As Long
    Dim I As Long, II As Long, tIcon As String
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    npcNum = MapNPC(Index).num
    
    ' Set the basic Y Value
    If NPC(npcNum).Sprite < 1 Or NPC(npcNum).Sprite > NumCharacters Then
        TextY = GetNPCTextY(Index) - 16
    Else
        ' Determine location for the text
        TextY = GetNPCTextY(Index) - (Tex_Character(NPC(npcNum).Sprite).Height / 4)
    
        If (Tex_Character(NPC(npcNum).Sprite).Height / 4) < 32 Then
            TextY = TextY - 16
        End If
    End If
    
    Select Case NPC(npcNum).Behavior
        Case NPC_BEHAVIOR_ATTACKONSIGHT
            Color = BrightRed
        Case NPC_BEHAVIOR_ATTACKWHENATTACKED
            Color = DarkGrey
        Case NPC_BEHAVIOR_GUARD
            Color = Magenta
        Case NPC_BEHAVIOR_QUEST
            Color = Yellow
    End Select
    
    Name = Trim$(NPC(npcNum).Name)
    
    If Len(Name) > 0 Then
        TextX = GetNPCTextX(Index) - (EngineGetTextWidth(Font_Default, Name) / 2)
        
        If Options.Levels = 1 And NPC(npcNum).Level > 0 Then
            Level = NPC(npcNum).Level
            
            If NPC(npcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(npcNum).Behavior = NPC_BEHAVIOR_ATTACKWHENATTACKED Then
                TextX = TextX + (EngineGetTextWidth(Font_Default, (Trim$(Level)))) + 4
            End If
        End If
        
        ' Draw Name
        Call RenderText(Font_Default, Name, TextX, TextY, Color)
    End If
    
    If Options.Levels = 1 And NPC(npcNum).Level > 0 Then
        If NPC(npcNum).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Or NPC(npcNum).Behavior = NPC_BEHAVIOR_ATTACKWHENATTACKED Then
            TextX = TextX - (EngineGetTextWidth(Font_Default, (Trim$(Level)))) - 4
            
            If NPC(npcNum).Level = GetPlayerLevel(MyIndex) Then
                Color = D3DColorARGB(255, 255, 255, 0)
            ElseIf NPC(npcNum).Level > GetPlayerLevel(MyIndex) Then
                Difference = NPC(npcNum).Level - GetPlayerLevel(MyIndex)
                If Difference > 10 Then Difference = 10
                Color = D3DColorARGB(255, 255, 255 - (255 * (Difference / 10)), 0)
            ElseIf NPC(npcNum).Level < GetPlayerLevel(MyIndex) Then
                Difference = GetPlayerLevel(MyIndex) - NPC(npcNum).Level
                If Difference > 10 Then Difference = 10
                Color = D3DColorARGB(255, 255 - (255 * (Difference / 10)), 255, 0)
            End If
            
            Call RenderText(Font_Default, Level, TextX, TextY, Color)
        End If
    End If
    
    If Len(Trim$(NPC(npcNum).title)) > 0 And Options.Titles = 1 Then
        TextX = GetNPCTextX(Index) + EngineGetTextWidth(Font_Default, Trim$(NPC(npcNum).title) / 2)
        
        ' Move it up
        TextY = TextY - 12
        
        ' Draw title
        Call RenderText(Font_Default, Trim$(NPC(npcNum).title), TextX, TextY, Color)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawNPCName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function DrawMapAttributes()
    Dim X As Long
    Dim Y As Long
    Dim tX As Long
    Dim tY As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Map.OptAttributes.Value Or frmMain.chkShowAttributes.Value Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    With Map.Tile(X, Y)
                        tX = ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5)
                        tY = ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                Call RenderText(Font_Default, "B", tX, tY, BrightRed)
                            Case TILE_TYPE_WARP
                                Call RenderText(Font_Default, "W", tX, tY, BrightBlue)
                            Case TILE_TYPE_ITEM
                                Call RenderText(Font_Default, "I", tX, tY, White)
                            Case TILE_TYPE_NPCAVOID
                                Call RenderText(Font_Default, "N", tX, tY, White)
                            Case TILE_TYPE_RESOURCE
                                Call RenderText(Font_Default, "R", tX, tY, Green)
                            Case TILE_TYPE_NPCSPAWN
                                Call RenderText(Font_Default, "S", tX, tY, Yellow)
                            Case TILE_TYPE_SHOP
                                Call RenderText(Font_Default, "S", tX, tY, BrightBlue)
                            Case TILE_TYPE_BANK
                                Call RenderText(Font_Default, "B", tX, tY, BrightBlue)
                            Case TILE_TYPE_HEAL
                                Call RenderText(Font_Default, "H", tX, tY, BrightGreen)
                            Case TILE_TYPE_TRAP
                                Call RenderText(Font_Default, "T", tX, tY, BrightRed)
                            Case TILE_TYPE_SLIDE
                                Call RenderText(Font_Default, "S", tX, tY, BrightCyan)
                            Case TILE_TYPE_CHECKPOINT
                                Call RenderText(Font_Default, "C", tX, tY, BrightGreen)
                            Case TILE_TYPE_SOUND
                                Call RenderText(Font_Default, "S", tX, tY, Orange)
                        End Select
                    End With
                End If
            Next
        Next
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "DrawMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub DrawActionMsg(ByVal Index As Long)
    Dim X As Long, Y As Long, I As Long, time As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Does it exist
    If ActionMsg(Index).Timer = 0 Then Exit Sub

    With ActionMsg(Index)

        If .WaitTimer > 0 Then
            If .WaitTimer > timeGetTime Then

                Exit Sub

            End If
        End If

        If .Alpha <= 0 Then
            Call ClearActionMsg(Index)

            Exit Sub

        End If

        ' Check if we should be seeing it
        If .Timer + 100 < timeGetTime Then
            .Alpha = .Alpha - 2.5
        End If

    End With

    ' How long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            time = 1500

            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            time = 1500
        
            If ActionMsg(Index).Y > 0 Then
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.4)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                X = ActionMsg(Index).X + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
                Y = ActionMsg(Index).Y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.4)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            time = 3000
            
            ' This will kill any action screen messages that there in the system
            For I = Action_HighIndex To 1 Step -1
                If ActionMsg(I).Type = ACTIONMSG_SCREEN Then
                    If Not I = Index Then
                        Call ClearActionMsg(Index)
                        Index = I
                    End If
                End If
            Next
            
            X = (frmMain.picScreen.Width \ 2) - ((Len(Trim$(ActionMsg(Index).Message)) \ 2) * 8)
            Y = 425
    End Select
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)

    If timeGetTime < ActionMsg(Index).Timer + time Then
        Call RenderText(Font_Default, ActionMsg(Index).Message, X, Y, ActionMsg(Index).Color, ActionMsg(Index).Alpha)
    Else
        ClearActionMsg Index
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub AddText(ByVal Msg As String, ByVal Color As Long)
    Dim S As String

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' No message just exit
    If Msg = vbNullString Then Exit Sub
    
    S = vbNewLine & Msg
    
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    
    If Color < Orange Then
        frmMain.txtChat.SelColor = QBColor(Color)
    Else
        frmMain.txtChat.SelColor = Color
    End If
    
    frmMain.txtChat.SelText = S
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SetPing()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    PingToDraw = Ping

    Select Case Ping
        Case -1
            PingToDraw = "Synchronizing"
        Case 0 To 10
            PingToDraw = "Local"
    End Select
    Exit Sub
    
ErrorHandler:
    HandleError "SetPing", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawEventName(ByVal Index As Long)
    Dim TextX As Long
    Dim TextY As Long
    Dim Color As Long
    Dim Name As String

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Don't draw while in the map editor
    If InMapEditor Then Exit Sub

    Color = White

    Name = Trim$(Map.MapEvents(Index).Name)
    
    ' Calc pos
    TextX = ConvertMapX(Map.MapEvents(Index).X * PIC_X) + Map.MapEvents(Index).xOffset + (PIC_X \ 2) + 1 - (EngineGetTextWidth(Font_Default, (Trim$(Name))) / 2)
    
    If Map.MapEvents(Index).GraphicType = 0 Then
        TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
    ElseIf Map.MapEvents(Index).GraphicType = 1 Then
        If Map.MapEvents(Index).GraphicNum < 1 Or Map.MapEvents(Index).GraphicNum > NumCharacters Then
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
        Else
            ' Determine location for text
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - (Tex_Character(Map.MapEvents(Index).GraphicNum).Height / 4) + 16
        End If
    ElseIf Map.MapEvents(Index).GraphicType = 2 Then
        If Map.MapEvents(Index).GraphicY2 > 0 Then
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - ((Map.MapEvents(Index).GraphicY2 - Map.MapEvents(Index).GraphicY) * 32) + 16
        Else
            TextY = ConvertMapY(Map.MapEvents(Index).Y * PIC_Y) + Map.MapEvents(Index).yOffset - 32 + 16
        End If
    End If

    ' Draw name
    RenderText Font_Default, Name, TextX, TextY, Color
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawEventName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
    Dim theArray() As String, X As Long, Y As Long, I As Long, MaxWidth As Long, X2 As Long, Y2 As Long, Color As Long
        
    With ChatBubble(Index)
        If .Alpha <= 0 Then
            Call ClearChatBubble(Index)
            Exit Sub
        End If
        
        ' Check if we should be seeing it
        If .Timer + 5000 < timeGetTime Then
            .Alpha = .Alpha - 7.5
        End If

        If .TargetType = TARGET_TYPE_PLAYER Then
            ' It's a player
            If GetPlayerMap(.Target) = GetPlayerMap(MyIndex) Then
                ' It's on our map - get co-ords
                X = ConvertMapX((Player(.Target).X * 32) + TempPlayer(.Target).xOffset) + 16
                Y = ConvertMapY((Player(.Target).Y * 32) + TempPlayer(.Target).yOffset) - 40
            End If
        ElseIf .TargetType = TARGET_TYPE_NPC Then
            ' It's on our map - get co-ords
            X = ConvertMapX((MapNPC(.Target).X * 32) + MapNPC(.Target).xOffset) + 16
            Y = ConvertMapY((MapNPC(.Target).Y * 32) + MapNPC(.Target).yOffset) - 40
        ElseIf .TargetType = TARGET_TYPE_EVENT Then
            X = ConvertMapX((Map.MapEvents(.Target).X * 32) + Map.MapEvents(.Target).xOffset) + 16
            Y = ConvertMapY((Map.MapEvents(.Target).Y * 32) + Map.MapEvents(.Target).yOffset) - 40
        End If
        
        ' Word wrap the text
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
                
        ' Find max width
        Dim Size As Long
        
        Size = UBound(theArray)
        
        For I = 1 To Size
            If EngineGetTextWidth(Font_Default, theArray(I)) > MaxWidth Then MaxWidth = EngineGetTextWidth(Font_Default, theArray(I))
        Next
                
        ' Calculate the new position
        X2 = X - (MaxWidth \ 2)
        Y2 = Y - (UBound(theArray) * 12)
                
        ' Render bubble - top left
        RenderTexture Tex_ChatBubble, X2 - 9, Y2 - 5, 0, 0, 9, 5, 9, 5, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Top right
        RenderTexture Tex_ChatBubble, X2 + MaxWidth, Y2 - 5, 119, 0, 9, 5, 9, 5, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Top
        RenderTexture Tex_ChatBubble, X2, Y2 - 5, 10, 0, MaxWidth, 5, 5, 5, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Bottom left
        RenderTexture Tex_ChatBubble, X2 - 9, Y, 0, 19, 9, 6, 9, 6, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Bottom right
        RenderTexture Tex_ChatBubble, X2 + MaxWidth, Y, 119, 19, 9, 6, 9, 6, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Bottom - left half
        RenderTexture Tex_ChatBubble, X2, Y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Bottom - right half
        RenderTexture Tex_ChatBubble, X2 + (MaxWidth \ 2) + 6, Y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Left
        RenderTexture Tex_ChatBubble, X2 - 9, Y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Right
        RenderTexture Tex_ChatBubble, X2 + MaxWidth, Y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Center
        RenderTexture Tex_ChatBubble, X2, Y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        
        ' Little pointy bit
        RenderTexture Tex_ChatBubble, X - 5, Y, 58, 19, 11, 11, 11, 11, D3DColorARGB(ChatBubble(Index).Alpha, 255, 255, 255)
        Size = UBound(theArray)
        ' Render each line centralised
        For I = 1 To Size
            RenderText Font_Georgia, theArray(I), X - (EngineGetTextWidth(Font_Default, theArray(I)) / 2), Y2, DarkBrown, .Alpha
            Y2 = Y2 + 12
        Next
        
        ' Check if it's timed out - close it if so
        If .Alpha = 0 Then
            .active = False
        End If
    End With
End Sub

Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
    Dim lineCount As Long, I As Long, Size As Long, lastSpace As Long, B As Long
    Dim textLen As Long
    
    textLen = Len(text)
    
    ' Too small of text
    If textLen < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If
    
    ' Default values
    B = 1
    lastSpace = 1
    Size = 0
    
    For I = 1 To textLen
        ' If it's a space, store it
        Select Case Mid$(text, I, 1)
            Case " ": lastSpace = I
            Case "_": lastSpace = I
            Case "-": lastSpace = I
        End Select
        
        ' Add up the size
        Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(text, I, 1)))
        
        ' Check for too large of a size
        If Size > MaxLineLen Then
            ' Check if the last space was too far back
            If I - lastSpace > 12 Then
                ' Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, (I - 1) - B))
                B = I - 1
                Size = 0
            Else
                ' Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, lastSpace - B))
                B = lastSpace + 1
                
                ' Count all the words we ignored (the ones that weren't printed, but are before "i")
                Size = EngineGetTextWidth(Font_Default, Mid$(text, lastSpace, I - lastSpace))
            End If
        End If
        
        ' Remainder
        If I = Len(text) Then
            If B <> I Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, B, I)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal text As String, ByVal MaxLineLen As Integer) As String
    Dim TempSplit() As String
    Dim TSLoop As Long
    Dim lastSpace As Long
    Dim Size As Long
    Dim I As Long
    Dim B As Long

    ' Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    ' Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)

    Size = UBound(TempSplit)
    For TSLoop = 0 To Size
        ' Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        ' Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        ' Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            ' Loop through all the characters
            Dim size2 As Long
            size2 = Len(TempSplit(TSLoop))
            For I = 1 To size2
            
                ' If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), I, 1)
                    Case " ": lastSpace = I
                    Case "_": lastSpace = I
                    Case "-": lastSpace = I
                End Select
    
                'Add up the size
                Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), I, 1)))
 
                'Check for too large of a size
                If Size > MaxLineLen Then
                    'Check if the last space was too far back
                    If I - lastSpace > 12 Then
                        ' Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (I - 1) - B)) & vbNewLine
                        B = I - 1
                        Size = 0
                    Else
                        ' Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)) & vbNewLine
                        B = lastSpace + 1
                        
                        ' Count all the words we ignored (the ones that weren't printed, but are before "i")
                        Size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, I - lastSpace))
                    End If
                End If
                
                ' This handles the remainder
                If I = Len(TempSplit(TSLoop)) Then
                    If B <> I Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), B, I)
                    End If
                End If
            Next I
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

Public Function GetPlayerTextX(ByVal Index As Long) As Long
    GetPlayerTextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + TempPlayer(Index).xOffset + (PIC_X \ 2) + 1
End Function

Public Function GetPlayerTextY(ByVal Index As Long) As Long
    GetPlayerTextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + TempPlayer(Index).yOffset + (PIC_Y \ 2)
End Function

Public Function GetNPCTextX(ByVal Index As Long) As Long
    GetNPCTextX = ConvertMapX(MapNPC(Index).X * PIC_X) + MapNPC(Index).xOffset + (PIC_X / 2) + 1
End Function

Public Function GetNPCTextY(ByVal Index As Long) As Long
    GetNPCTextY = ConvertMapY(MapNPC(Index).Y * PIC_Y) + MapNPC(Index).yOffset + (PIC_Y / 2)
End Function
