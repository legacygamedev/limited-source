Attribute VB_Name = "modText"
Option Explicit

' ******************************************
' **               rootSource               **
' ** GDI text drawing                     **
' ******************************************

' Text declares
Private Declare Function CreateFont Lib "gdi32.dll" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32.dll" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32.dll" Alias "TextOutA" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long

' Font variables
Private Const FONT_NAME As String = "fixedsys"
Private Const FONT_SIZE As Byte = 18

' Text variables
Public TexthDC As Long
Private MainFont As Long

' Draw map name location
Public DrawMapNameX As Single
Public DrawMapNameY As Single
Public DrawMapNameColor As Long

' Used to check if text needs to be drawn
Public BFPS As Boolean ' FPS
Public BLoc As Boolean ' map, player, and mouse location

' Quote character variable
Public vbQuote As String ' container for "
' Public Const vbQuote As String = """"

Public Sub InitFont()
    Call SetFont(MainFont, FONT_NAME, FONT_SIZE)
End Sub

Private Sub SetFont(ByRef Font As Long, ByVal FontName As String, ByVal FontSize As Byte)
    Font = CreateFont(FontSize, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, FontName)
End Sub

' Draw text onto buffer
Public Sub DrawText(ByVal hDC As Long, ByVal x, ByVal y, ByVal Text As String, Color As Long)
    
    ' Selects an object into the specified DC.
    ' The new object replaces the previous object of the same type.
    Call SelectObject(hDC, MainFont)
    
    ' Sets the background mix mode of the specified DC.
    Call SetBkMode(hDC, vbTransparent)
    
    ' color of text drop shadow
    Call SetTextColor(hDC, RGB(0, 0, 0))
    
    ' draw text shadow with offset
    Call TextOut(hDC, x + 2, y + 2, Text, Len(Text))
    Call TextOut(hDC, x + 1, y + 1, Text, Len(Text))
    
    ' draw text with color
    Call SetTextColor(hDC, Color)
    Call TextOut(hDC, x, y, Text, Len(Text))
End Sub

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim x2, y2 As Long
Dim insight As Boolean
    
                insight = False
                Select Case Player(Index).map
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
                
    If insight Then
    
    ' Check access level to determine color
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(Brown)
            Case 1
                Color = QBColor(Gray)
            Case 2
                Color = QBColor(Green)
            Case 3
                Color = QBColor(Blue)
            Case 4
                Color = QBColor(Yellow)
            Case 5
                Color = QBColor(White)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If

    ' Determine location for text
    TextX = (x2 + GetPlayerX(Index)) * PIC_X + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerName(Index)) / 2) * 8) + StaticX
    TextY = (y2 + GetPlayerY(Index)) * PIC_Y + Player(Index).YOffset - (DDS_Sprite(GetPlayerSprite(Index)).SurfDescription.lHeight) + 16 + StaticY
    
    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(Index), Color)

    End If
End Sub

Public Sub DrawPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
Dim x2, y2 As Long
Dim insight As Boolean
     
If Player(Index).Guild = "" Then Exit Sub
    
                insight = False
                Select Case Player(Index).map
                    Case map(5).Right
                        insight = True
                        x2 = MAX_MAPX + 1
                        y2 = 0
                    Case map(5).Left
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = 0
                    Case map(5).Up
                        insight = True
                        x2 = 0
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(5).Down
                        insight = True
                        x2 = 0
                        y2 = MAX_MAPY + 1
                    Case map(4).Up 'north west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down 'south west
                        insight = True
                        x2 = -1 * (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case map(6).Up 'north east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = -1 * (MAX_MAPY + 1)
                    Case map(4).Down ' south east
                        insight = True
                        x2 = (MAX_MAPX + 1)
                        y2 = (MAX_MAPY + 1)
                    Case Player(MyIndex).map
                        insight = True
                        x2 = 0
                        y2 = 0
                End Select
                
    If insight Then
    
    ' Check access level to determine color
        Select Case Player(Index).GuildAccess
            Case 0
                Color = QBColor(Brown)
            Case 1
                Color = QBColor(Gray)
            Case 2
                Color = QBColor(Green)
        End Select

    ' Determine location for text
    TextX = (x2 + GetPlayerX(Index)) * PIC_X + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerName(Index)) / 2) * 8) + StaticX
    TextY = (y2 + GetPlayerY(Index)) * PIC_Y - 12 + Player(Index).YOffset - (DDS_Sprite(GetPlayerSprite(Index)).SurfDescription.lHeight) + 16 + StaticY
    
    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, Player(Index).Guild, Color)

    End If
End Sub

Public Function DrawMapAttributes()
    Dim x As Long
    Dim y As Long
   
    For x = 0 To MAX_MAPX
        For y = 0 To MAX_MAPY

            With map(5).Tile(x, y)

                Select Case .Type
                   
                    Case TILE_TYPE_BLOCKED
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "B", QBColor(BrightRed)
                       
                    Case TILE_TYPE_WARP
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "W", QBColor(BrightBlue)
                   
                    Case TILE_TYPE_ITEM
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "I", QBColor(White)
                   
                    Case TILE_TYPE_NPCAVOID
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "N", QBColor(White)
                   
                    Case TILE_TYPE_KEY
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "K", QBColor(White)
                   
                    Case TILE_TYPE_KEYOPEN
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "O", QBColor(White)
                   
                    Case TILE_TYPE_HEAL
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "H", QBColor(White)
                                           
                    Case TILE_TYPE_KILL
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "K", QBColor(White)
                   
                    Case TILE_TYPE_DOOR
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "D", QBColor(White)
                   
                    Case TILE_TYPE_SIGN
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "S", QBColor(White)

                    Case TILE_TYPE_MSG
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "M", QBColor(White)
                   
                    Case TILE_TYPE_SPRITE
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "SP", QBColor(White)
                   
                    Case TILE_TYPE_NPCSPAWN
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "NS", QBColor(White)
                   
                    Case TILE_TYPE_NUDGE
                        DrawText TexthDC, ((x * PIC_X) - 4) + (PIC_X * 0.5), ((y * PIC_Y) - 7) + (PIC_Y * 0.5), "NU", QBColor(White)

                End Select

            End With

        Next
    Next
   
End Function

Public Function getSize(ByVal DC As Long, ByVal Text As String) As TextSize
    Dim lngReturn As Long
    Dim typSize As TextSize

    lngReturn = GetTextExtentPoint32(DC, Text, Len(Text), typSize)

    getSize = typSize
End Function

