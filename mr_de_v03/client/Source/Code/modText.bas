Attribute VB_Name = "modText"
Option Explicit
Public Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As textSize) As Long
Public Type textSize
    Width As Long
    Height As Long
End Type

Public Sub SetFont(ByVal Font As String, ByVal Size As Byte)
    GameFont = CreateFont(Size, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, Font)
End Sub

Public Function getTextWidth(ByRef Text As String) As Long
Dim lngReturn As Long
Dim typSize As textSize
   
    lngReturn = GetTextExtentPoint32(TexthDC, Text, Len(Text), typSize)
    getTextWidth = typSize.Width
End Function

Public Sub DrawText(ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal Text As String, Color As Long)
    Call SelectObject(hdc, GameFont)
    Call SetBkMode(hdc, vbTransparent)
    Call SetTextColor(hdc, RGB(0, 0, 0))
    'Call TextOut(hdc, X + 1, y + 1, Text, Len(Text))

    'Call TextOut(hdc, X + 2, y + 2, Text, Len(Text))
    Call TextOut(hdc, X + 2, Y + 0, Text, Len(Text))
    Call TextOut(hdc, X + 1, Y + 0, Text, Len(Text))
    Call TextOut(hdc, X + 0, Y + 1, Text, Len(Text))
    Call TextOut(hdc, X + 0, Y + 2, Text, Len(Text))

    'Call TextOut(hdc, X - 1, y - 1, Text, Len(Text))

    'Call TextOut(hdc, X - 2, y - 2, Text, Len(Text))
    Call TextOut(hdc, X - 2, Y - 0, Text, Len(Text))
    Call TextOut(hdc, X - 1, Y - 0, Text, Len(Text))
    Call TextOut(hdc, X - 0, Y - 1, Text, Len(Text))
    Call TextOut(hdc, X - 0, Y - 2, Text, Len(Text))

    'Call TextOut(hdc, X + 1, y - 1, Text, Len(Text))

    'Call TextOut(hdc, X + 2, y - 2, Text, Len(Text))
    Call TextOut(hdc, X + 2, Y - 0, Text, Len(Text))
    Call TextOut(hdc, X + 1, Y - 0, Text, Len(Text))
    Call TextOut(hdc, X + 0, Y - 1, Text, Len(Text))
    Call TextOut(hdc, X + 0, Y - 2, Text, Len(Text))

    'Call TextOut(hdc, X - 1, y + 1, Text, Len(Text))

    'Call TextOut(hdc, X - 2, y + 2, Text, Len(Text))
    Call TextOut(hdc, X - 2, Y + 0, Text, Len(Text))
    Call TextOut(hdc, X - 1, Y + 0, Text, Len(Text))
    Call TextOut(hdc, X - 0, Y + 1, Text, Len(Text))
    Call TextOut(hdc, X - 0, Y + 2, Text, Len(Text))

    Call SetTextColor(hdc, Color)
    Call TextOut(hdc, X, Y, Text, Len(Text))
End Sub

'Public Sub AddText(ByVal Msg As String, ByVal Color As Integer)
'Dim s As String
'
'    s = vbNewLine & Msg
'    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text)
'    frmMainGame.txtChat.SelColor = QBColor(Color)
'    frmMainGame.txtChat.SelText = s
'    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text) - 1
'End Sub

Public Sub AddText(ByRef Msg As String, ByVal Color As Integer)
Dim s As String, tempS As String
Dim sLength As Long, sStart As Long, sEnd As Long

    s = vbNewLine & Msg
    ' Let's get the length of the current chat text for setting the current selText
    sLength = Len(frmMainGame.txtChat.Text)
    If sLength = 0 Then sLength = 1
    
    frmMainGame.txtChat.SelStart = sLength
    frmMainGame.txtChat.SelColor = QBColor(Color)
    frmMainGame.txtChat.SelText = s
    
    If ShowItemLinks Then
        ' Checks if we have both a start and end to a itemBlock
        tempS = frmMainGame.txtChat.Text
        sStart = InStr(sLength, tempS, "[")
        sEnd = InStr(sLength, tempS, "]")
        If sStart > 0 Then
            If sEnd > 0 Then
                ' This makes sure that the start of the itemBlock is before the ending
                If sStart < sEnd Then
                    CheckForCommand tempS, sLength
                End If
            End If
        End If
    End If
    
    frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text) - 1
    
End Sub

Public Sub CheckForCommand(ByRef Msg As String, ByVal start As Long)
Dim ItemName As String, ItemNum As String
Dim sLength As Long, sStart As Long, sEnd As Long

    ' Checks to make sure we have some valid shit
    If start <= 0 Then Exit Sub
    If Len(Msg) <= 0 Then Exit Sub
    
    ' Now that we have a full command lets get the length inside the command
    sStart = InStr(start, Msg, "[") + 1
    sEnd = InStr(start, Msg, "]")
    sLength = sEnd - sStart
    
    ' Get the item name
    ItemName = Mid$(Msg, sStart, sLength)
    
    ' Get the number, if there's no number it's not a valid item
    ItemNum = GetItemNumFromName(ItemName)
    If ItemNum > 0 Then
        frmMainGame.txtChat.Find "[" & ItemName & "]", start
        frmMainGame.Hypertext.InsertHyperlink ItemNum
        frmMainGame.txtChat.SelStart = Len(frmMainGame.txtChat.Text) - 1
    End If
    
    sLength = InStr(start, Msg, "[" & ItemName & "]") + Len("[" & ItemName & "]")
    sStart = InStr(sLength, Msg, "[")
    sEnd = InStr(sLength, Msg, "]")
    If sStart > 0 Then
        If sEnd > 0 Then
            If sStart < sEnd Then
                CheckForCommand Msg, sLength
            End If
        End If
    End If
End Sub

Public Function GetItemNumFromName(ByVal ItemName As String) As Long
Dim i As Long
    For i = 1 To MAX_ITEMS
        If Trim$(Item(i).Name) = ItemName Then
            GetItemNumFromName = i
            Exit Function
        End If
    Next
End Function

Public Sub CacheItem(ByVal ItemNum As Long)
    If Trim$(Item(ItemNum).Name) <> vbNullString Then
        ItemReq(ItemNum) = ItemRequirement(ItemNum)
        ItemDesc(ItemNum) = ItemDescription(ItemNum)
    Else
        ItemReq(ItemNum) = "No data available"
        ItemDesc(ItemNum) = "No data available"
    End If
End Sub

Public Function ItemRequirement(ByVal ItemNum As Long) As String
Dim i As Long
Dim tempS As String

    tempS = "Requirements" & vbNewLine
    If Item(ItemNum).LevelReq > 0 Then
        tempS = tempS & "Level " & CStr(Item(ItemNum).LevelReq) & vbNewLine
    End If

    For i = 0 To MAX_CLASSES
        If Item(ItemNum).ClassReq And (2 ^ i) Then
            tempS = tempS & Trim$(Class(i).Name) & vbNewLine
        End If
    Next
    
    For i = 1 To Stats.Stat_Count
        If Item(ItemNum).StatReq(i) > 0 Then
            tempS = tempS & CStr(Item(ItemNum).StatReq(i)) & " " & StatName(i) & vbNewLine
        End If
    Next
    
    ItemRequirement = tempS
End Function

Public Function ItemDescription(ByVal ItemNum As Long) As String
Dim i As Long
Dim tempS As String
      
    tempS = "Description" & vbNewLine
    For i = 1 To Vitals.Vital_Count
        If Item(ItemNum).ModVital(i) > 0 Then
            tempS = tempS & "+" & CStr(Item(ItemNum).ModVital(i)) & " " & VitalName(i) & vbNewLine
        ElseIf Item(ItemNum).ModVital(i) < 0 Then
            tempS = tempS & CStr(Item(ItemNum).ModVital(i)) & " " & VitalName(i) & vbNewLine
        End If
    Next

    For i = 1 To Stats.Stat_Count
        If Item(ItemNum).ModStat(i) > 0 Then
            tempS = tempS & "+" & CStr(Item(ItemNum).ModStat(i)) & " " & StatName(i) & vbNewLine
        ElseIf Item(ItemNum).ModStat(i) < 0 Then
            tempS = tempS & CStr(Item(ItemNum).ModStat(i)) & " " & StatName(i) & vbNewLine
        End If
    Next
    
    ' If the item is bound
    If Item(ItemNum).Bound Then
        tempS = tempS & vbNewLine & BindName(Item(ItemNum).Bound) & vbNewLine
    End If
    
    ItemDescription = tempS
End Function

Public Sub BltAttributes()
Dim X As Long
Dim Y As Long
    ' Blit out attribs if in editor
    If InEditor Then
        If frmMapEditor.optAttribs.Value Or frmMapEditor.optNpcs.Value Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.Bottom
                    If IsValidMapPoint(X, Y) Then
                        Select Case Map.Tile(X, Y).Type
                            Case TILE_TYPE_BLOCKED
                                DrawText TexthDC, ConvertMapX(X) + 8, ConvertMapY(Y) + 8, "B", QBColor(BrightRed)
                            Case TILE_TYPE_WARP
                                DrawText TexthDC, ConvertMapX(X) + 8, ConvertMapY(Y) + 8, "W", QBColor(BrightBlue)
                            Case TILE_TYPE_ITEM
                                DrawText TexthDC, ConvertMapX(X) + 8, ConvertMapY(Y) + 8, "I", QBColor(Yellow)
                            Case TILE_TYPE_NPCAVOID
                                DrawText TexthDC, ConvertMapX(X) + 8, ConvertMapY(Y) + 8, "N", QBColor(White)
                            Case TILE_TYPE_KEY
                                DrawText TexthDC, ConvertMapX(X) + 8, ConvertMapY(Y) + 8, "D", QBColor(Pink)
                            Case TILE_TYPE_KEYOPEN
                                DrawText TexthDC, ConvertMapX(X) + 8, ConvertMapY(Y) + 8, "S", QBColor(BrightGreen)
                            Case TILE_TYPE_MOBSPAWN
                                DrawText TexthDC, ConvertMapX(X) + 8, ConvertMapY(Y) + 8, Map.Tile(X, Y).Data1, QBColor(Yellow)
                        End Select
                    End If
                Next
            DoEvents
            Next
            
        End If
    End If
End Sub
