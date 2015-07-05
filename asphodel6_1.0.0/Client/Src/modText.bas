Attribute VB_Name = "modText"
Option Explicit

' ------------------------------------------
' --               Asphodel               --
' ------------------------------------------

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Text As String, Color As Long)
    
    ' Keep the text from going off screen
    If X < 0 Then X = 0
    If X > ((MAX_MAPX + 1) * PIC_X) - (Len(Text) * FONT_WIDTH) Then X = ((MAX_MAPX + 1) * PIC_X) - (Len(Text) * FONT_WIDTH)
    
    If Y < 0 Then Y = 0
    If Y > ((MAX_MAPY + 1) * PIC_Y) - FONT_HEIGHT Then X = ((MAX_MAPY + 1) * PIC_Y) - FONT_HEIGHT
    
    ' Selects an object into the specified DC.
    ' The new object replaces the previous object of the same type.
    SelectObject hdc, GameFont
    
    ' Sets the background mix mode of the specified DC.
    SetBkMode hdc, vbTransparent
    
    ' color of text drop shadow
    SetTextColor hdc, 0
    
    ' draw with border
    TextOut hdc, X - 1, Y, Text, Len(Text)
    TextOut hdc, X + 1, Y, Text, Len(Text)
    TextOut hdc, X, Y - 1, Text, Len(Text)
    TextOut hdc, X, Y + 1, Text, Len(Text)
    
    ' draw text with color
    SetTextColor hdc, Color
    TextOut hdc, X, Y, Text, Len(Text)
    
End Sub

Public Sub DrawPlayerGuildNames()
Dim TextX As Long
Dim TextY As Long
Dim Use_Color As Long
Dim Index As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Index = i
                
                If LenB(Trim$(Player(Index).GuildName)) > 0 Then
                    ' Check access level to determine color
                    Select Case Player(Index).GuildRank
                        Case 1
                            Use_Color = ColorTable(Color.green)
                        Case 2
                            Use_Color = ColorTable(Color.Grey)
                        Case 3
                            Use_Color = ColorTable(Color.Brown)
                        Case 4
                            Use_Color = ColorTable(Color.White)
                    End Select
                    
                    ' Determine location for text
                    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + (PIC_X \ 2) - ((Len("[" & Player(Index).GuildName & "]") / 2) * FONT_WIDTH)
                    TextY = (GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - Sprite_Offset) - (FONT_HEIGHT + 4)
                    
                    If GetPlayerSprite(Index) > -1 Then
                        If GetPlayerSprite(Index) <= TOTAL_SPRITES Then
                            If Sprite_Size(GetPlayerSprite(Index)).SizeY > PIC_Y Then
                                TextY = TextY - (Sprite_Size(GetPlayerSprite(Index)).SizeY - PIC_Y)
                            End If
                        End If
                    End If
                    
                    ' Draw name
                    DrawText TexthDC, TextX, TextY, "[" & Player(Index).GuildName & "]", Use_Color
                End If
            End If
        End If
    Next
    
End Sub

Public Sub DrawPlayerNames()
Dim TextX As Long
Dim TextY As Long
Dim Use_Color As Long
Dim Index As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Index = i
                
                ' Check access level to determine color
                If GetPlayerPK(Index) = NO Then
                    Select Case GetPlayerAccess(Index)
                        Case 0
                            Use_Color = ColorTable(Color.Brown)
                        Case 1
                            Use_Color = ColorTable(Color.DarkGrey)
                        Case 2
                            Use_Color = ColorTable(Color.Cyan)
                        Case 3
                            Use_Color = ColorTable(Color.blue)
                        Case 4
                            Use_Color = ColorTable(Color.Pink)
                    End Select
                Else
                    Use_Color = ColorTable(Color.BrightRed)
                End If
                
                ' Determine location for text
                TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerName(Index)) / 2) * FONT_WIDTH)
                TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - (PIC_Y \ 2) - Sprite_Offset
                
                If GetPlayerSprite(Index) > -1 Then
                    If GetPlayerSprite(Index) <= TOTAL_SPRITES Then
                        If Sprite_Size(GetPlayerSprite(Index)).SizeY > PIC_Y Then
                            TextY = TextY - (Sprite_Size(GetPlayerSprite(Index)).SizeY - PIC_Y)
                        End If
                    End If
                End If
                
                ' Draw name
                DrawText TexthDC, TextX, TextY, GetPlayerName(Index), Use_Color
            End If
        End If
    Next

End Sub

Public Sub DrawNpcNames()
Dim TextX As Long
Dim TextY As Long
Dim i As Long

    For i = 1 To UBound(MapNpc)
        If MapNpc(i).Num > 0 Then
            
            ' Determine location for text
            TextX = GetMapNpcX(i) * PIC_X + MapNpc(i).XOffset + (PIC_X \ 2) - ((Len(Trim$(GetMapNpcName(i))) / 2) * FONT_WIDTH)
            TextY = GetMapNpcY(i) * PIC_Y + MapNpc(i).YOffset - (PIC_Y \ 2) - Sprite_Offset
            
            If GetMapNpcSprite(i) > -1 Then
                If GetMapNpcSprite(i) <= TOTAL_SPRITES Then
                    If Sprite_Size(GetMapNpcSprite(i)).SizeY > PIC_Y Then
                        TextY = TextY - (Sprite_Size(GetMapNpcSprite(i)).SizeY - PIC_Y)
                    End If
                End If
            End If
            
            ' Draw name
            DrawText TexthDC, TextX, TextY, GetMapNpcName(i), ColorTable(Color.White)
            
        End If
    Next

End Sub
