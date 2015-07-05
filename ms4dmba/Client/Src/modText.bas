Attribute VB_Name = "modText"
Option Explicit

' ******************************************
' **            Mirage Source 4           **
' ** GDI text drawing                     **
' ******************************************

' Text declares
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal i As Long, ByVal u As Long, ByVal s As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal f As String) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

' Used to set a font for GDI text drawing
Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

' GDI text drawing onto buffer
Public Sub DrawText(ByVal hdc As Long, ByVal X, ByVal Y, ByVal Text As String, Color As Long)
    
    ' Selects an object into the specified DC.
    ' The new object replaces the previous object of the same type.
    Call SelectObject(hdc, GameFont)
    
    ' Sets the background mix mode of the specified DC.
    Call SetBkMode(hdc, vbTransparent)
    
    ' color of text drop shadow
    Call SetTextColor(hdc, RGB(0, 0, 0))
    
    ' draw with offset
    Call TextOut(hdc, X + 2, Y + 2, Text, Len(Text))
    Call TextOut(hdc, X + 1, Y + 1, Text, Len(Text))
    
    ' draw text with color
    Call SetTextColor(hdc, Color)
    Call TextOut(hdc, X, Y, Text, Len(Text))
End Sub

Public Sub DrawPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    
    ' Check access level to determine color
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(Brown)
            Case 1
                Color = QBColor(DarkGrey)
            Case 2
                Color = QBColor(Cyan)
            Case 3
                Color = QBColor(Blue)
            Case 4
                Color = QBColor(Pink)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If

    ' Determine location for text
    TextX = ConvertMapX(GetPlayerX(Index) * PIC_X) + Player(Index).XOffset + (PIC_X \ 2) - ((Len(GetPlayerName(Index)) / 2) * 8)
    TextY = ConvertMapY(GetPlayerY(Index) * PIC_Y) + Player(Index).YOffset - (PIC_Y \ 2) - 4
    
    ' Draw name
    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(Index), Color)
End Sub

Public Function BltMapAttributes()
    Dim X As Long
    Dim Y As Long

    If frmMirage.optAttribs.Value Then
        For X = TileView.Left To TileView.Right
            For Y = TileView.Top To TileView.Bottom
                If IsValidMapPoint(X, Y) Then
                    With Map.Tile(X, Y)
                        Select Case .Type
                            Case TILE_TYPE_BLOCKED
                                DrawText TexthDC, ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5), ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5), "B", QBColor(BrightRed)
                            
                            Case TILE_TYPE_WARP
                                DrawText TexthDC, ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5), ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5), "W", QBColor(BrightBlue)
                        
                            Case TILE_TYPE_ITEM
                                DrawText TexthDC, ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5), ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5), "I", QBColor(White)
                        
                            Case TILE_TYPE_NPCAVOID
                                DrawText TexthDC, ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5), ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5), "N", QBColor(White)
                        
                            Case TILE_TYPE_KEY
                                DrawText TexthDC, ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5), ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5), "K", QBColor(White)
                        
                            Case TILE_TYPE_KEYOPEN
                                DrawText TexthDC, ((ConvertMapX(X * PIC_X)) - 4) + (PIC_X * 0.5), ((ConvertMapY(Y * PIC_Y)) - 7) + (PIC_Y * 0.5), "O", QBColor(White)
                        
                        End Select
                    End With
                End If
            Next
        Next
    End If
    
End Function

