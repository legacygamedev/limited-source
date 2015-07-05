Attribute VB_Name = "modText"
Option Explicit

' text color pointers
Public Const Black As Byte = 0
Public Const Blue As Byte = 1
Public Const Green As Byte = 2
Public Const Cyan As Byte = 3
Public Const Red As Byte = 4
Public Const Magenta As Byte = 5
Public Const Brown As Byte = 6
Public Const Grey As Byte = 7
Public Const DarkGrey As Byte = 8
Public Const BrightBlue As Byte = 9
Public Const BrightGreen As Byte = 10
Public Const BrightCyan As Byte = 11
Public Const BrightRed As Byte = 12
Public Const Pink As Byte = 13
Public Const Yellow As Byte = 14
Public Const White As Byte = 15
Public Const DarkBrown As Byte = 16

Public Const SayColor As Byte = White
Public Const GlobalColor As Byte = BrightBlue
Public Const BroadcastColor As Byte = White
Public Const TellColor As Byte = BrightGreen
Public Const EmoteColor As Byte = BrightCyan
Public Const AdminColor As Byte = BrightCyan
Public Const HelpColor As Byte = BrightBlue
Public Const WhoColor As Byte = BrightBlue
Public Const JoinLeftColor As Byte = DarkGrey
Public Const NpcColor As Byte = Brown
Public Const AlertColor As Byte = Red
Public Const NewMapColor As Byte = BrightBlue

' Stuffs
Public Type POINTAPI
    x As Long
    y As Long
End Type

Private Type CharVA
    Vertex(0 To 3) As TLVERTEX
End Type

Private Type VFH
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
    Texture As Direct3DTexture8
    RowPitch As Integer
    RowFactor As Single
    ColFactor As Single
    CharHeight As Byte
    TextureSize As POINTAPI
End Type

' Chat Buffer
Public ChatVA() As TLVERTEX
Public ChatVAS() As TLVERTEX

Public Const ChatTextBufferSize As Integer = 200
Public ChatBufferChunk As Single

'Text buffer
Public Type ChatTextBuffer
    Text As String
    Color As Long
End Type

'Chat vertex buffer information
Public ChatArrayUbound As Long
Public ChatVB As Direct3DVertexBuffer8
Public ChatVBS As Direct3DVertexBuffer8
Public ChatTextBuffer(1 To ChatTextBufferSize) As ChatTextBuffer

Public Font_Default As CustomFont
Public Font_Georgia As CustomFont

Public Sub DrawPlayerName(ByVal Index As Long)
Dim textX As Long, textY As Long, Text As String, textSize As Long, colour As Long
    
    Text = Trim$(GetPlayerName(Index))
    textSize = EngineGetTextWidth(Font_Default, Text)
    
    ' get the colour
    If GetPlayerAccess(Index) > 0 Then
        colour = Yellow
    Else
        colour = White
    End If
    
    textX = Player(Index).x * PIC_X + Player(Index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = Player(Index).y * PIC_Y + Player(Index).yOffset - 32
    
    If GetPlayerSprite(Index) >= 1 And GetPlayerSprite(Index) <= Count_Char Then
        textY = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - (D3DT_TEXTURE(Tex_Char(GetPlayerSprite(Index))).height / 4) + 12
    End If
    
    Call RenderText(Font_Default, Text, ConvertMapX(textX), ConvertMapY(textY), colour)
End Sub

Public Sub DrawNpcName(ByVal Index As Long)
Dim textX As Long, textY As Long, Text As String, textSize As Long, npcNum As Long, colour As Long
    
    npcNum = MapNpc(Index).num
    Text = Trim$(Npc(npcNum).Name)
    textSize = EngineGetTextWidth(Font_Default, Text)
    
    If Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKONSIGHT Or Npc(npcNum).Behaviour = NPC_BEHAVIOUR_ATTACKWHENATTACKED Then
        ' get the colour
        If Npc(npcNum).Level <= GetPlayerLevel(MyIndex) - 3 Then
            colour = Grey
        ElseIf Npc(npcNum).Level <= GetPlayerLevel(MyIndex) - 2 Then
            colour = Green
        ElseIf Npc(npcNum).Level > GetPlayerLevel(MyIndex) Then
            colour = Red
        Else
            colour = White
        End If
    Else
        colour = White
    End If
    
    textX = MapNpc(Index).x * PIC_X + MapNpc(Index).xOffset + (PIC_X \ 2) - (textSize \ 2)
    textY = MapNpc(Index).y * PIC_Y + MapNpc(Index).yOffset - 32
    
    If Npc(npcNum).sprite >= 1 And Npc(npcNum).sprite <= Count_Char Then
        textY = MapNpc(Index).y * PIC_Y + MapNpc(Index).yOffset - (D3DT_TEXTURE(Tex_Char(Npc(npcNum).sprite)).height / 4) + 12
    End If
    
    Call RenderText(Font_Default, Text, ConvertMapX(textX), ConvertMapY(textY), colour)
End Sub

Public Sub RenderText(ByRef UseFont As CustomFont, ByVal Text As String, ByVal x As Long, ByVal y As Long, ByVal Color As Long, Optional ByVal alpha As Long = 255, Optional Shadow As Boolean = True)
Dim TempVA(0 To 3)  As TLVERTEX
Dim TempVAS(0 To 3) As TLVERTEX
Dim TempStr() As String
Dim Count As Integer
Dim Ascii() As Byte
Dim Row As Integer
Dim u As Single
Dim v As Single
Dim i As Long
Dim j As Long
Dim KeyPhrase As Byte
Dim TempColor As Long
Dim ResetColor As Byte
Dim srcRECT As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim yOffset As Single

    ' set the color
    Color = dx8Colour(Color, alpha)

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    'Get the text into arrays (split by vbCrLf)
    TempStr = Split(Text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = Color
    
    'Set the texture
    D3DDevice8.SetTexture 0, UseFont.Texture
    CurrentTexture = -1
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(TempStr)
        If Len(TempStr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
            'Convert the characters to the ascii value
            Ascii() = StrConv(TempStr(i), vbFromUnicode)
            
            'Loop through the characters
            For j = 1 To Len(TempStr(i))
                'Copy from the cached vertex array to the temp vertex array
                Call CopyMemory(TempVA(0), UseFont.HeaderInfo.CharVA(Ascii(j - 1)).Vertex(0), FVF_Size * 4)
                
                'Set up the verticies
                TempVA(0).x = x + Count
                TempVA(0).y = y + yOffset
                TempVA(1).x = TempVA(1).x + x + Count
                TempVA(1).y = TempVA(0).y
                TempVA(2).x = TempVA(0).x
                TempVA(2).y = TempVA(2).y + TempVA(0).y
                TempVA(3).x = TempVA(1).x
                TempVA(3).y = TempVA(2).y
                
                'Set the colors
                TempVA(0).Color = TempColor
                TempVA(1).Color = TempColor
                TempVA(2).Color = TempColor
                TempVA(3).Color = TempColor
                
                'Draw the verticies
                Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, TempVA(0), FVF_Size)
                
                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(Ascii(j - 1))
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
            Next j
        End If
    Next i
End Sub

Sub EngineInitFontTextures()
    'Check if we have the device
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then Exit Sub

    ' FONT DEFAULT
    Set Font_Default.Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, App.path & Path_Font & "texdefault.png", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, RGB(255, 0, 255), ByVal 0, ByVal 0)
    Font_Default.TextureSize.x = 256
    Font_Default.TextureSize.y = 256
    
    ' Georgia
    Set Font_Georgia.Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, App.path & Path_Font & "georgia.png", 256, 256, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, RGB(255, 0, 255), ByVal 0, ByVal 0)
    Font_Georgia.TextureSize.x = 256
    Font_Georgia.TextureSize.y = 256
End Sub

Sub LoadFontHeader(ByRef theFont As CustomFont, ByVal FileName As String)
Dim FileNum As Byte
Dim LoopChar As Long
Dim Row As Single
Dim u As Single
Dim v As Single

    'Load the header information
    FileNum = FreeFile
    Open App.path & Path_Font & FileName For Binary As #FileNum
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
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).RHW = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).x = 0
            .Vertex(0).y = 0
            .Vertex(0).z = 0
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).RHW = 1
            .Vertex(1).tu = u + theFont.ColFactor
            .Vertex(1).tv = v
            .Vertex(1).x = theFont.HeaderInfo.CellWidth
            .Vertex(1).y = 0
            .Vertex(1).z = 0
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).RHW = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + theFont.RowFactor
            .Vertex(2).x = 0
            .Vertex(2).y = theFont.HeaderInfo.CellHeight
            .Vertex(2).z = 0
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).RHW = 1
            .Vertex(3).tu = u + theFont.ColFactor
            .Vertex(3).tv = v + theFont.RowFactor
            .Vertex(3).x = theFont.HeaderInfo.CellWidth
            .Vertex(3).y = theFont.HeaderInfo.CellHeight
            .Vertex(3).z = 0
        End With
    Next LoopChar
End Sub

Sub EngineInitFontSettings()
    LoadFontHeader Font_Default, "texdefault.dat"
    LoadFontHeader Font_Georgia, "georgia.dat"
End Sub

Public Function dx8Colour(ByVal colourNum As Long, ByVal alpha As Long) As Long
    Select Case colourNum
        Case 0 ' Black
            dx8Colour = D3DColorARGB(alpha, 0, 0, 0)
        Case 1 ' Blue
            dx8Colour = D3DColorARGB(alpha, 16, 104, 237)
        Case 2 ' Green
            dx8Colour = D3DColorARGB(alpha, 119, 188, 84)
        Case 3 ' Cyan
            dx8Colour = D3DColorARGB(alpha, 16, 224, 237)
        Case 4 ' Red
            dx8Colour = D3DColorARGB(alpha, 201, 0, 0)
        Case 5 ' Magenta
            dx8Colour = D3DColorARGB(alpha, 255, 0, 255)
        Case 6 ' Brown
            dx8Colour = D3DColorARGB(alpha, 175, 149, 92)
        Case 7 ' Grey
            dx8Colour = D3DColorARGB(alpha, 192, 192, 192)
        Case 8 ' DarkGrey
            dx8Colour = D3DColorARGB(alpha, 128, 128, 128)
        Case 9 ' BrightBlue
            dx8Colour = D3DColorARGB(alpha, 126, 182, 240)
        Case 10 ' BrightGreen
            dx8Colour = D3DColorARGB(alpha, 126, 240, 137)
        Case 11 ' BrightCyan
            dx8Colour = D3DColorARGB(alpha, 157, 242, 242)
        Case 12 ' BrightRed
            dx8Colour = D3DColorARGB(alpha, 255, 0, 0)
        Case 13 ' Pink
            dx8Colour = D3DColorARGB(alpha, 255, 118, 221)
        Case 14 ' Yellow
            dx8Colour = D3DColorARGB(alpha, 255, 255, 0)
        Case 15 ' White
            dx8Colour = D3DColorARGB(alpha, 255, 255, 255)
        Case 16 ' dark brown
            dx8Colour = D3DColorARGB(alpha, 98, 84, 52)
    End Select
End Function

Public Function EngineGetTextWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
Dim LoopI As Integer

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For LoopI = 1 To Len(Text)
        EngineGetTextWidth = EngineGetTextWidth + UseFont.HeaderInfo.CharWidth(Asc(Mid$(Text, LoopI, 1)))
    Next LoopI

End Function

Sub DrawActionMsg(ByVal Index As Integer)
Dim x As Long, y As Long, i As Long, Time As Long
Dim LenMsg As Long, alpha As Long

    If ActionMsg(Index).message = vbNullString Then Exit Sub

    ' how long we want each message to appear
    Select Case ActionMsg(Index).Type
        Case ACTIONMSG_STATIC
            Time = 1500
            
            LenMsg = EngineGetTextWidth(Font_Default, Trim$(ActionMsg(Index).message))

            If ActionMsg(Index).y > 0 Then
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - (LenMsg / 2)
                y = ActionMsg(Index).y + PIC_Y
            Else
                x = ActionMsg(Index).x + Int(PIC_X \ 2) - (LenMsg / 2)
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
                y = ActionMsg(Index).y - Int(PIC_Y \ 2) + 18 + (ActionMsg(Index).Scroll * 0.001)
                ActionMsg(Index).Scroll = ActionMsg(Index).Scroll + 1
            End If
            
            ActionMsg(Index).alpha = ActionMsg(Index).alpha - 5
            If ActionMsg(Index).alpha <= 0 Then ClearActionMsg Index: Exit Sub

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
    
            x = (400) - ((EngineGetTextWidth(Font_Default, Trim$(ActionMsg(Index).message)) \ 2))
            y = 24

    End Select
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    If ActionMsg(Index).Created > 0 Then
        RenderText Font_Default, ActionMsg(Index).message, x, y, ActionMsg(Index).Color, ActionMsg(Index).alpha
    End If

End Sub

Public Function DrawMapAttributes()
    Dim x As Long
    Dim y As Long
    Dim tX As Long
    Dim tY As Long

    If frmEditor_Map.optAttribs.value Then
        For x = TileView.left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    With Map.Tile(x, y)
                        tX = ((ConvertMapX(x * PIC_X)) - 4) + (PIC_X * 0.5)
                        tY = ((ConvertMapY(y * PIC_Y)) - 7) + (PIC_Y * 0.5)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                RenderText Font_Default, "B", tX, tY, BrightRed
                            Case TILE_TYPE_WARP
                                RenderText Font_Default, "W", tX, tY, BrightBlue
                            Case TILE_TYPE_ITEM
                                RenderText Font_Default, "I", tX, tY, White
                            Case TILE_TYPE_NPCAVOID
                                RenderText Font_Default, "N", tX, tY, White
                            Case TILE_TYPE_KEY
                                RenderText Font_Default, "K", tX, tY, White
                            Case TILE_TYPE_KEYOPEN
                                RenderText Font_Default, "O", tX, tY, White
                            Case TILE_TYPE_RESOURCE
                                RenderText Font_Default, "R", tX, tY, Green
                            Case TILE_TYPE_DOOR
                                RenderText Font_Default, "D", tX, tY, Brown
                            Case TILE_TYPE_NPCSPAWN
                                RenderText Font_Default, "S", tX, tY, Yellow
                            Case TILE_TYPE_SHOP
                                RenderText Font_Default, "S", tX, tY, BrightBlue
                            Case TILE_TYPE_SLIDE
                                RenderText Font_Default, "S", tX, tY, Pink
                            Case TILE_TYPE_CHAT
                                RenderText Font_Default, "C", tX, tY, Blue
                        End Select
                    End With
                End If
            Next
        Next
    End If

End Function

' Chat Box
Public Sub RenderChatTextBuffer()
Dim srcRECT As RECT
Dim v2 As D3DVECTOR2
Dim v3 As D3DVECTOR2
Dim i As Long

    'Clear the LastTexture, letting the rest of the engine know that the texture needs to be changed for next rect render
    D3DDevice8.SetTexture 0, Font_Default.Texture
    CurrentTexture = -1

    If ChatArrayUbound > 0 Then
        D3DDevice8.SetStreamSource 0, ChatVBS, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
        D3DDevice8.SetStreamSource 0, ChatVB, FVF_Size
        D3DDevice8.DrawPrimitive D3DPT_TRIANGLELIST, 0, (ChatArrayUbound + 1) \ 3
    End If
    
End Sub

Public Sub UpdateChatArray()
Dim Chunk As Integer
Dim Count As Integer
Dim LoopC As Byte
Dim Ascii As Byte
Dim Row As Long
Dim Pos As Long
Dim u As Single
Dim v As Single
Dim x As Single
Dim y As Single
Dim y2 As Single
Dim i As Long
Dim j As Long
Dim size As Integer
Dim KeyPhrase As Byte
Dim ResetColor As Byte
Dim TempColor As Long
Dim yOffset As Long

    ' set the offset of each line
    yOffset = 14

    'Set the position
    If ChatBufferChunk <= 1 Then ChatBufferChunk = 1
    
    Chunk = ChatScroll
    
    'Get the number of characters in all the visible buffer
    size = 0
    
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        size = size + Len(ChatTextBuffer(LoopC).Text)
    Next
    
    size = size - j
    ChatArrayUbound = size * 6 - 1
    If ChatArrayUbound < 0 Then Exit Sub
    ReDim ChatVA(0 To ChatArrayUbound) 'Size our array to fix the 6 verticies of each character
    ReDim ChatVAS(0 To ChatArrayUbound)
    
    'Set the base position
    x = GUIWindow(GUI_CHAT).x + ChatOffsetX
    y = GUIWindow(GUI_CHAT).y + ChatOffsetY

    'Loop through each buffer string
    For LoopC = (Chunk * ChatBufferChunk) - (8 - 1) To Chunk * ChatBufferChunk
        If LoopC > ChatTextBufferSize Then Exit For
        If ChatBufferChunk * Chunk > ChatTextBufferSize Then ChatBufferChunk = ChatBufferChunk - 1
        
        'Set the temp color
        TempColor = ChatTextBuffer(LoopC).Color
        
        'Set the Y position to be used
        y2 = y - (LoopC * yOffset) + (Chunk * ChatBufferChunk * yOffset) - 32
        
        'Loop through each line if there are line breaks (vbCrLf)
        Count = 0   'Counts the offset value we are on
        If LenB(ChatTextBuffer(LoopC).Text) <> 0 Then  'Dont bother with empty strings
            
            'Loop through the characters
            For j = 1 To Len(ChatTextBuffer(LoopC).Text)
            
                'Convert the character to the ascii value
                Ascii = Asc(Mid$(ChatTextBuffer(LoopC).Text, j, 1))
                
                'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
                Row = (Ascii - Font_Default.HeaderInfo.BaseCharOffset) \ Font_Default.RowPitch
                u = ((Ascii - Font_Default.HeaderInfo.BaseCharOffset) - (Row * Font_Default.RowPitch)) * Font_Default.ColFactor
                v = Row * Font_Default.RowFactor

                ' ****** Rectangle | Top Left ******
                With ChatVA(0 + (6 * Pos))
                    .Color = TempColor
                    .x = (x) + Count
                    .y = (y2)
                    .tu = u
                    .tv = v
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Left ******
                With ChatVA(1 + (6 * Pos))
                    .Color = TempColor
                    .x = (x) + Count
                    .y = (y2) + Font_Default.HeaderInfo.CellHeight
                    .tu = u
                    .tv = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                
                ' ****** Rectangle | Bottom Right ******
                With ChatVA(2 + (6 * Pos))
                    .Color = TempColor
                    .x = (x) + Count + Font_Default.HeaderInfo.CellWidth
                    .y = (y2) + Font_Default.HeaderInfo.CellHeight
                    .tu = u + Font_Default.ColFactor
                    .tv = v + Font_Default.RowFactor
                    .RHW = 1
                End With
                
                
                'Triangle 2 (only one new vertice is needed)
                ChatVA(3 + (6 * Pos)) = ChatVA(0 + (6 * Pos)) 'Top-left corner
                
                ' ****** Rectangle | Top Right ******
                With ChatVA(4 + (6 * Pos))
                    .Color = TempColor
                    .x = (x) + Count + Font_Default.HeaderInfo.CellWidth
                    .y = (y2)
                    .tu = u + Font_Default.ColFactor
                    .tv = v
                    .RHW = 1
                End With

                ChatVA(5 + (6 * Pos)) = ChatVA(2 + (6 * Pos))

                'Update the character we are on
                Pos = Pos + 1

                'Shift over the the position to render the next character
                Count = Count + Font_Default.HeaderInfo.CharWidth(Ascii)
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = ChatTextBuffer(LoopC).Color
                End If
            Next
        End If
    Next LoopC
        
    If Not D3DDevice8 Is Nothing Then   'Make sure the D3DDevice exists - this will only return false if we received messages before it had time to load
        Set ChatVBS = D3DDevice8.CreateVertexBuffer(FVF_Size * Pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVBS, 0, FVF_Size * Pos * 6, 0, ChatVAS(0)
        Set ChatVB = D3DDevice8.CreateVertexBuffer(FVF_Size * Pos * 6, 0, FVF, D3DPOOL_MANAGED)
        D3DVertexBuffer8SetData ChatVB, 0, FVF_Size * Pos * 6, 0, ChatVA(0)
    End If
    Erase ChatVAS()
    Erase ChatVA()
    
End Sub

Public Sub AddText(ByVal Text As String, ByVal tColor As Long, Optional ByVal alpha As Long = 255)
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim size As Long
Dim i As Long
Dim b As Long
Dim Color As Long

    Color = dx8Colour(tColor, alpha)

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbCrLf)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        size = 0
        b = 1
        lastSpace = 1
        
        'Loop through all the characters
        For i = 1 To Len(TempSplit(TSLoop))
        
            'If it is a space, store it so we can easily break at it
            Select Case Mid$(TempSplit(TSLoop), i, 1)
                Case " ": lastSpace = i
                Case "_": lastSpace = i
                Case "-": lastSpace = i
            End Select
            
            'Add up the size
            size = size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
            
            'Check for too large of a size
            If size > ChatWidth Then
                
                'Check if the last space was too far back
                If i - lastSpace > 10 Then
                
                    'Too far away to the last space, so break at the last character
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)), Color
                    b = i - 1
                    size = 0
                Else
                    'Break at the last space to preserve the word
                    AddToChatTextBuffer_Overflow Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)), Color
                    b = lastSpace + 1
                    'Count all the words we ignored (the ones that weren't printed, but are before "i")
                    size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                End If
            End If
            
            'This handles the remainder
            If i = Len(TempSplit(TSLoop)) Then
                If b <> i Then AddToChatTextBuffer_Overflow Mid$(TempSplit(TSLoop), b, i), Color
            End If
        Next i
    Next TSLoop
    
    'Only update if we have set up the text (that way we can add to the buffer before it is even made)
    If Font_Default.RowPitch = 0 Then Exit Sub
    
    If ChatScroll > 8 Then ChatScroll = ChatScroll + 1

    'Update the array
    UpdateChatArray
End Sub

Private Sub AddToChatTextBuffer_Overflow(ByVal Text As String, ByVal Color As Long)
Dim LoopC As Long

    'Move all other text up
    For LoopC = (ChatTextBufferSize - 1) To 1 Step -1
        ChatTextBuffer(LoopC + 1) = ChatTextBuffer(LoopC)
    Next LoopC
    
    'Set the values
    ChatTextBuffer(1).Text = Text
    ChatTextBuffer(1).Color = Color
    
    ' set the total chat lines
    totalChatLines = totalChatLines + 1
    If totalChatLines > ChatTextBufferSize - 1 Then totalChatLines = ChatTextBufferSize - 1
End Sub

Public Sub WordWrap_Array(ByVal Text As String, ByVal MaxLineLen As Long, ByRef theArray() As String)
Dim lineCount As Long, i As Long, size As Long, lastSpace As Long, b As Long
    
    'Too small of text
    If Len(Text) < 2 Then
        ReDim theArray(1 To 1) As String
        theArray(1) = Text
        Exit Sub
    End If
    
    ' default values
    b = 1
    lastSpace = 1
    size = 0
    
    For i = 1 To Len(Text)
        ' if it's a space, store it
        Select Case Mid$(Text, i, 1)
            Case " ": lastSpace = i
            Case "_": lastSpace = i
            Case "-": lastSpace = i
        End Select
        
        'Add up the size
        size = size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(Text, i, 1)))
        
        'Check for too large of a size
        If size > MaxLineLen Then
            'Check if the last space was too far back
            If i - lastSpace > 12 Then
                'Too far away to the last space, so break at the last character
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, b, (i - 1) - b))
                b = i - 1
                size = 0
            Else
                'Break at the last space to preserve the word
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = Trim$(Mid$(Text, b, lastSpace - b))
                b = lastSpace + 1
                
                'Count all the words we ignored (the ones that weren't printed, but are before "i")
                size = EngineGetTextWidth(Font_Default, Mid$(Text, lastSpace, i - lastSpace))
            End If
        End If
        
        ' Remainder
        If i = Len(Text) Then
            If b <> i Then
                lineCount = lineCount + 1
                ReDim Preserve theArray(1 To lineCount) As String
                theArray(lineCount) = theArray(lineCount) & Mid$(Text, b, i)
            End If
        End If
    Next
End Sub

Public Function WordWrap(ByVal Text As String, ByVal MaxLineLen As Integer) As String
Dim TempSplit() As String
Dim TSLoop As Long
Dim lastSpace As Long
Dim size As Long
Dim i As Long
Dim b As Long

    'Too small of text
    If Len(Text) < 2 Then
        WordWrap = Text
        Exit Function
    End If

    'Check if there are any line breaks - if so, we will support them
    TempSplit = Split(Text, vbNewLine)
    
    For TSLoop = 0 To UBound(TempSplit)
    
        'Clear the values for the new line
        size = 0
        b = 1
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
                size = size + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(TempSplit(TSLoop), i, 1)))
 
                'Check for too large of a size
                If size > MaxLineLen Then
                    'Check if the last space was too far back
                    If i - lastSpace > 12 Then
                        'Too far away to the last space, so break at the last character
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, (i - 1) - b)) & vbNewLine
                        b = i - 1
                        size = 0
                    Else
                        'Break at the last space to preserve the word
                        WordWrap = WordWrap & Trim$(Mid$(TempSplit(TSLoop), b, lastSpace - b)) & vbNewLine
                        b = lastSpace + 1
                        
                        'Count all the words we ignored (the ones that weren't printed, but are before "i")
                        size = EngineGetTextWidth(Font_Default, Mid$(TempSplit(TSLoop), lastSpace, i - lastSpace))
                    End If
                End If
                
                'This handles the remainder
                If i = Len(TempSplit(TSLoop)) Then
                    If b <> i Then
                        WordWrap = WordWrap & Mid$(TempSplit(TSLoop), b, i)
                    End If
                End If
            Next i
        Else
            WordWrap = WordWrap & TempSplit(TSLoop)
        End If
    Next TSLoop
End Function

Public Sub UpdateShowChatText()
Dim CHATOFFSET As Long, i As Long, x As Long

    CHATOFFSET = 52
    
    If EngineGetTextWidth(Font_Default, MyText) > GUIWindow(GUI_CHAT).Width - CHATOFFSET Then
        For i = Len(MyText) To 1 Step -1
            x = x + Font_Default.HeaderInfo.CharWidth(Asc(Mid$(MyText, i, 1)))
            If x > GUIWindow(GUI_CHAT).Width - CHATOFFSET Then
                RenderChatText = Right$(MyText, Len(MyText) - i + 1)
                Exit For
            End If
        Next
    Else
        RenderChatText = MyText
    End If
End Sub
