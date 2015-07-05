Attribute VB_Name = "modText"
Option Explicit
' Stuffs
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Public Type VFH
    BitmapWidth As Long
    BitmapHeight As Long
    CellWidth As Long
    CellHeight As Long
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

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal text As String, ByVal x As Long, ByVal y As Long, ByVal color As Long, Optional ByVal Alpha As Long = 0, Optional Shadow As Boolean = True)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempVAS(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRect As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim yOffset As Single

    ' set the color
    Alpha = 255 - Alpha
    color = dx8Colour(color, Alpha)
    
    'Check for valid text to render
    If LenB(text) = 0 Then Exit Sub
    
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = color
    
    'Set the texture
    Direct3D_Device.SetTexture 0, gTexture(UseFont.Texture.Texture).Texture
    'CurrentTexture = -1
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            
            'Loop through the characters
            For j = 1 To Len(TempStr(i))
                'Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_SIZE * 4)
                
                'Set up the verticies
                TempVA(0).x = x + count
                TempVA(0).y = y + yOffset
                TempVA(1).x = TempVA(1).x + x + count
                TempVA(1).y = TempVA(0).y
                TempVA(2).x = TempVA(0).x
                TempVA(2).y = TempVA(2).y + TempVA(0).y
                TempVA(3).x = TempVA(1).x
                TempVA(3).y = TempVA(2).y
                
                'Set the colors
                TempVA(0).color = TempColor
                TempVA(1).color = TempColor
                TempVA(2).color = TempColor
                TempVA(3).color = TempColor
                
                'Draw the verticies
                Call Direct3D_Device.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0)))
                
                'Shift over the the position to render the next character
                count = count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = color
                End If
            Next j
        End If
    Next i
End Sub

Sub EngineInitFontTextures()
    ' FONT DEFAULT
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
Sub UnloadFont(Font As CustomFont)
    Font.Texture.Texture = 0
End Sub


Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal filename As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single


    'Load the header information
    FileNum = FreeFile
    Open App.Path & FONT_PATH & filename For Binary As #FileNum
        Get #FileNum, , theFont.HeaderInfo
    Close #FileNum
    
    'Calculate some common values
    theFont.CharHeight = theFont.HeaderInfo.CellHeight - 4
    theFont.RowPitch = theFont.HeaderInfo.BitmapWidth \ theFont.HeaderInfo.CellWidth
    theFont.ColFactor = theFont.HeaderInfo.CellWidth / theFont.HeaderInfo.BitmapWidth
    theFont.RowFactor = theFont.HeaderInfo.CellHeight / theFont.HeaderInfo.BitmapHeight
    
    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255
        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - theFont.HeaderInfo.BaseCharOffset) \ theFont.RowPitch
        u = ((LoopChar - theFont.HeaderInfo.BaseCharOffset) - (Row * theFont.RowPitch)) * theFont.ColFactor
        v = Row * theFont.RowFactor
        
        'Set the verticies
        With theFont.HeaderInfo.CharVA(LoopChar)
            .Vertex(0).color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).TU = u
            .Vertex(0).TV = v
            .Vertex(0).x = 0
            .Vertex(0).y = 0
            .Vertex(0).Z = 0
            .Vertex(1).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).TU = u + theFont.ColFactor
            .Vertex(1).TV = v
            .Vertex(1).x = theFont.HeaderInfo.CellWidth
            .Vertex(1).y = 0
            .Vertex(1).Z = 0
            .Vertex(2).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).TU = u
            .Vertex(2).TV = v + theFont.RowFactor
            .Vertex(2).x = 0
            .Vertex(2).y = theFont.HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            .Vertex(3).color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).TU = u + theFont.ColFactor
            .Vertex(3).TV = v + theFont.RowFactor
            .Vertex(3).x = theFont.HeaderInfo.CellWidth
            .Vertex(3).y = theFont.HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
    Next LoopChar
End Sub

Sub EngineInitFontSettings()
    LoadFontHeader Font_Default, "texdefault.dat"
    LoadFontHeader Font_Georgia, "georgia.dat"
End Sub
Public Function dx8Colour(ByVal colourNum As Long, ByVal Alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(Alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(Alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(Alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(Alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(Alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(Alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(Alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(Alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(Alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(Alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(Alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(Alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(Alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(Alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(Alpha, 98, 84, 52)
        Case 17 'Orange
            dx8Colour = D3DColorARGB(Alpha, 255, 96, 0)
    End Select
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text
    If LenB(text) = 0 Then Exit Function
    
    'Loop through the text
    For LoopI = 1 To Len(text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(text, LoopI, 1)))
    Next LoopI

End Function

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Check access level
    If GetPlayerPK(Index) = NO Then

        Select Case GetPlayerAccess(Index)
            Case 0
                color = Orange
            Case 1
                color = White
            Case 2
                color = Cyan
            Case 3
                color = BrightGreen
            Case 4
                color = Yellow
        End Select

    Else
        color = BrightRed
    End If

    name = Trim$(Player(Index).name)
    ' calc pos
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(name))) / 2)
    If GetPlayerSprite(Index) < 1 Or GetPlayerSprite(Index) > NumCharacters Then
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).yOffset - (Tex_Character(GetPlayerSprite(Index)).Height / 4) + 16
    End If

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    RenderText Font_Default, name, TextX, TextY, color, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawPlayerName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim name As String
Dim npcNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    npcNum = MapNpc(Index).num

    Select Case Npc(npcNum).Behaviour
        Case NPC_BEHAVIOUR_ATTACKONSIGHT
            color = BrightRed
        Case NPC_BEHAVIOUR_ATTACKWHENATTACKED
            color = Yellow
        Case NPC_BEHAVIOUR_GUARD
            color = Grey
        Case Else
            color = BrightGreen
    End Select

    name = Trim$(Npc(npcNum).name)
    TextX = ConvertMapX(MapNpc(Index).x * PIC_X) + MapNpc(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(name))) / 2)
    If Npc(npcNum).Sprite < 1 Or Npc(npcNum).Sprite > NumCharacters Then
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - 16
    Else
        ' Determine location for text
        TextY = ConvertMapY(MapNpc(Index).y * PIC_Y) + MapNpc(Index).yOffset - (Tex_Character(Npc(npcNum).Sprite).Height / 4) + 16
    End If

    ' Draw name
    'Call DrawText(TexthDC, TextX, TextY, Name, Color)
    RenderText Font_Default, name, TextX, TextY, color, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawNpcName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function DrawMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tx As Long
    Dim ty As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If frmEditor_Map.optAttribs.value Then
        For x = TileView.Left To TileView.Right
            For y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(x, y) Then
                    With Map.Tile(x, y)
                        tx = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        ty = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText Font_Default, "B", tx, ty, BrightRed, 0
                            Case TILE_TYPE_WARP
                                RenderText Font_Default, "W", tx, ty, BrightBlue, 0
                            Case TILE_TYPE_ITEM
                                RenderText Font_Default, "I", tx, ty, White, 0
                            Case TILE_TYPE_NPCAVOID
                                RenderText Font_Default, "N", tx, ty, White, 0
                            Case TILE_TYPE_KEY
                                RenderText Font_Default, "K", tx, ty, White, 0
                            Case TILE_TYPE_KEYOPEN
                                RenderText Font_Default, "O", tx, ty, White, 0
                            Case TILE_TYPE_RESOURCE
                                RenderText Font_Default, "B", tx, ty, Green, 0
                            Case TILE_TYPE_DOOR
                                RenderText Font_Default, "D", tx, ty, Brown, 0
                            Case TILE_TYPE_NPCSPAWN
                                RenderText Font_Default, "S", tx, ty, Yellow, 0
                            Case TILE_TYPE_SHOP
                                RenderText Font_Default, "S", tx, ty, BrightBlue, 0
                            Case TILE_TYPE_BANK
                                RenderText Font_Default, "B", tx, ty, Blue, 0
                            Case TILE_TYPE_HEAL
                                RenderText Font_Default, "H", tx, ty, BrightGreen, 0
                            Case TILE_TYPE_TRAP
                                RenderText Font_Default, "T", tx, ty, BrightRed, 0
                            Case TILE_TYPE_SLIDE
                                RenderText Font_Default, "S", tx, ty, BrightCyan, 0
                            Case TILE_TYPE_SOUND
                                RenderText Font_Default, "S", tx, ty, Orange, 0
                        End Select
                    End With
                End If
            Next
        Next
    End If

    ' Error handler
    Exit Function
errorhandler:
    HandleError "DrawMapAttributes", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Sub DrawActionMsg(ByVal Index As Long)
    Dim x As Long, y As Long, i As Long, Time As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' does it exist
    If ActionMsg(Index).Created = 0 Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500

            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18
            End If

        Case ACTIONMSG_SCROLL
            Time = 1500
        
            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) - 2 - (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.6)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If

        Case ACTIONMSG_SCREEN
            Time = 3000

            ' This will kill any action screen messages that there in the system
            For i = MAX_BYTE To 1 Step -1
                If ActionMsg(i).Type = ACTIONMSG_SCREEN Then
                    If i <> Index Then
                        ClearActionMsg Index
                        Index = i
                    End If
                End If
            Next
            x = (frmMain.picScreen.Width \ 2) - ((Len(Trim$(ActionMsg(Index).message)) \ 2) * 8)
            y = 425

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If GetTickCount < ActionMsg(Index).Created + Time Then
        RenderText Font_Default, ActionMsg(Index).message, x, y, ActionMsg(Index).color, 0
    Else
        ClearActionMsg Index
    End If

    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawActionMsg", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Function getWidth(Font As CustomFont, ByVal text As String) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    getWidth = EngineGetTextWidth(Font, text)
    
    ' Error handler
    Exit Function
errorhandler:
    HandleError "getWidth", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Function
End Function

Public Sub AddText(ByVal Msg As String, ByVal color As Integer)
Dim s As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    s = vbNewLine & Msg
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = QBColor(color)
    frmMain.txtChat.SelText = s
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "AddText", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawEventName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim color As Long
Dim name As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    If InMapEditor Then Exit Sub

    color = White

    name = Trim$(Map.MapEvents(Index).name)
    
    ' calc pos
    TextX = ConvertMapX(Map.MapEvents(Index).x * PIC_X) + Map.MapEvents(Index).xOffset + (PIC_X \ 2) - (getWidth(Font_Default, (Trim$(name))) / 2)
    If Map.MapEvents(Index).GraphicType = 0 Then
        TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
    ElseIf Map.MapEvents(Index).GraphicType = 1 Then
        If Map.MapEvents(Index).GraphicNum < 1 Or Map.MapEvents(Index).GraphicNum > NumCharacters Then
            TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - 16
        Else
            ' Determine location for text
            TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - (Tex_Character(Map.MapEvents(Index).GraphicNum).Height / 4) + 16
        End If
    ElseIf Map.MapEvents(Index).GraphicType = 2 Then
        If Map.MapEvents(Index).GraphicY2 > 0 Then
            TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - ((Map.MapEvents(Index).GraphicY2 - Map.MapEvents(Index).GraphicY) * 32) + 16
        Else
            TextY = ConvertMapY(Map.MapEvents(Index).y * PIC_Y) + Map.MapEvents(Index).yOffset - 32 + 16
        End If
    End If

    ' Draw name
    RenderText Font_Default, name, TextX, TextY, color, 0
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "DrawEventName", "modText", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
Dim theArray() As String, x As Long, y As Long, i As Long, MaxWidth As Long, x2 As Long, y2 As Long, colour As Long
    
    With chatBubble(Index)
        If .targetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' it's on our map - get co-ords
                x = ConvertMapX((Player(.target).x * 32) + Player(.target).xOffset) + 16
                y = ConvertMapY((Player(.target).y * 32) + Player(.target).yOffset) - 40
            End If
        ElseIf .targetType = TARGET_TYPE_NPC Then
            ' it's on our map - get co-ords
            x = ConvertMapX((MapNpc(.target).x * 32) + MapNpc(.target).xOffset) + 16
            y = ConvertMapY((MapNpc(.target).y * 32) + MapNpc(.target).yOffset) - 40
        ElseIf .targetType = TARGET_TYPE_EVENT Then
            x = ConvertMapX((Map.MapEvents(.target).x * 32) + Map.MapEvents(.target).xOffset) + 16
            y = ConvertMapY((Map.MapEvents(.target).y * 32) + Map.MapEvents(.target).yOffset) - 40
        End If
        
        ' word wrap the text
        WordWrap_Array .Msg, ChatBubbleWidth, theArray
                
        ' find max width
        For i = 1 To UBound(theArray)
            If EngineGetTextWidth(Font_Default, theArray(i)) > MaxWidth Then MaxWidth = EngineGetTextWidth(Font_Default, theArray(i))
        Next
                
        ' calculate the new position
        x2 = x - (MaxWidth \ 2)
        y2 = y - (UBound(theArray) * 12)
                
        ' render bubble - top left
        RenderTexture Tex_ChatBubble, x2 - 9, y2 - 5, 0, 0, 9, 5, 9, 5
        ' top right
        RenderTexture Tex_ChatBubble, x2 + MaxWidth, y2 - 5, 119, 0, 9, 5, 9, 5
        ' top
        RenderTexture Tex_ChatBubble, x2, y2 - 5, 10, 0, MaxWidth, 5, 5, 5
        ' bottom left
        RenderTexture Tex_ChatBubble, x2 - 9, y, 0, 19, 9, 6, 9, 6
        ' bottom right
        RenderTexture Tex_ChatBubble, x2 + MaxWidth, y, 119, 19, 9, 6, 9, 6
        ' bottom - left half
        RenderTexture Tex_ChatBubble, x2, y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' bottom - right half
        RenderTexture Tex_ChatBubble, x2 + (MaxWidth \ 2) + 6, y, 10, 19, (MaxWidth \ 2) - 5, 6, 9, 6
        ' left
        RenderTexture Tex_ChatBubble, x2 - 9, y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
        ' right
        RenderTexture Tex_ChatBubble, x2 + MaxWidth, y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
        ' center
        RenderTexture Tex_ChatBubble, x2, y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
        ' little pointy bit
        RenderTexture Tex_ChatBubble, x - 5, y, 58, 19, 11, 11, 11, 11
                
        ' render each line centralised
        For i = 1 To UBound(theArray)
            RenderText Font_Georgia, theArray(i), x - (EngineGetTextWidth(Font_Default, theArray(i)) / 2), y2, DarkBrown
            y2 = y2 + 12
        Next
        ' check if it's timed out - close it if so
        If .timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub

Public Sub WordWrap_Array(ByVal text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, i As Long, Size As Long, lastSpace As Long, B As Long
    
    'Too small of text
    If Len(text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = text
        Exit Sub
    End If
    
    ' default values
    B = 1
    lastSpace = 1
    Size = 0
    
    For i = 1 To Len(text)
        ' if it's a space, store it
        Select Case Mid$(text, i, 1)
            Case " ": lastSpace = i
            Case "_": lastSpace = i
            Case "-": lastSpace = i
        End Select
        
        'Add up the size
        Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(text, i, 1)))
        
        'Check for too large of a size
        If Size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, (i - 1) - B))
                B = i - 1
                Size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(text, B, lastSpace - B))
                B = lastSpace + 1
                
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                Size = EngineGetTextWidth(Font_Default, Mid$(text, lastSpace, i - lastSpace))
            End If
        End If
        
        ' Remainder
        If i = Len(text) Then
            If B <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(text, B, i)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal text As String, ByVal MaxLineLen As Integer) As String
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim Size As Long
Dim i As Long
Dim B As Long

    'Too small of text
    If Len(text) < 2 Then
        WordWrap = text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        Size = 0
        B = 1
        lastSpace = 1
        
        'Add back in the vbNewLines
        If TSLoop < UBound(TempSplit()) Then TempSplit(TSLoop) = TempSplit(TSLoop) & vbNewLine
        
        'Only check lines with a space
        If InStr(1, TempSplit(TSLoop), " ") Or InStr(1, TempSplit(TSLoop), "-") Or InStr(1, TempSplit(TSLoop), "_") Then
            
            'Loop through all the characters
            For i = 1 To Len(TempSplit(TSLoop))
            
                'If it is a space, store it so we can easily break at it
                Select Case Mid$(TempSplit(TSLoop), i, 1)
                    Case " ": lastSpace = i
                    Case "_": lastSpace = i
                    Case "-": lastSpace = i
                End Select
    
                'Add up the size
                Size = Size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
 
                'Check for too large of a size
                If Size > MaxLineLen Then
                    'Check if the last space was too far back
                    If i - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, (i - 1) - B)) & vbNewLine
                        B = i - 1
                        Size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), B, lastSpace - B)) & vbNewLine
                        B = lastSpace + 1
                        
                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        Size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                    End If
                End If
                
                'This handles the remainder
                If i = Len(TempSplit(TSLoop)) Then
                    If B <> i Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), B, i)
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function
