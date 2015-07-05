Attribute VB_Name = "modLoot"
Option Explicit

Function FindOpenMapItemSlot(ByVal MapNum As Long) As Long
Dim i As Long

    FindOpenMapItemSlot = 0
    
    For i = 1 To MAX_MAP_ITEMS
        If MapData(MapNum).MapItem(i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next
End Function

Sub SpawnItem(ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal DropTime As Long)
Dim i As Long

    ' Check for subscript out of range
    If ItemNum < 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
    If MapNum <= 0 Then Exit Sub
    If MapNum > MAX_MAPS Then Exit Sub
    
    ' Find open map item slot
    i = FindOpenMapItemSlot(MapNum)
    
    SpawnItemSlot i, ItemNum, ItemVal, MapNum, X, Y, DropTime
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Long, ByVal ItemVal As Long, ByVal MapNum As Long, ByVal X As Long, ByVal Y As Long, ByVal DropTime As Long)

    ' Check for subscript out of range
    If MapItemSlot <= 0 Then Exit Sub
    If MapItemSlot > MAX_MAP_ITEMS Then Exit Sub
    If ItemNum < 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
    If MapNum <= 0 Then Exit Sub
    If MapNum > MAX_MAPS Then Exit Sub
    
    MapData(MapNum).MapItem(MapItemSlot).Num = ItemNum
    MapData(MapNum).MapItem(MapItemSlot).Value = ItemVal
    MapData(MapNum).MapItem(MapItemSlot).X = X
    MapData(MapNum).MapItem(MapItemSlot).Y = Y
    MapData(MapNum).MapItem(MapItemSlot).DropTime = DropTime
        
    SendSpawnItem MapNum, MapItemSlot
End Sub

Sub SpawnAllMapsItems()
Dim i As Long
    
    For i = 1 To MAX_MAPS
        SpawnMapItems i
    Next
End Sub

Sub SpawnMapItems(ByVal MapNum As Long)
Dim X As Long
Dim Y As Long

    ' Check for subscript out of range
    If MapNum <= 0 Then Exit Sub
    If MapNum > MAX_MAPS Then Exit Sub
    
    ' Spawn what we have
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_ITEM Then
                SpawnItem Map(MapNum).Tile(X, Y).Data1, Map(MapNum).Tile(X, Y).Data2, MapNum, X, Y, 0
            End If
        Next
    Next
End Sub

Sub PlayerMapGetItem(ByVal Index As Long)
Dim i As Long
Dim n As Long
Dim MapNum As Long

Dim DeleteItem As Boolean
Dim ItemNum As Long
Dim ItemValue As Long

    If Not IsPlaying(Index) Then Exit Sub
    
    MapNum = Current_Map(Index)
    
    For i = MAX_MAP_ITEMS To 1 Step -1
        
        ItemNum = MapData(MapNum).MapItem(i).Num
        
        ' See if theres even an item here
        If ItemNum > 0 Then
            If ItemNum <= MAX_ITEMS Then
                ' Check if item is at the same location as the player
                If MapData(MapNum).MapItem(i).X = Current_X(Index) Then
                    If MapData(MapNum).MapItem(i).Y = Current_Y(Index) Then
                        
                        ' Find open slot
                        n = FindOpenInvSlot(Index, ItemNum)
                        
                        ' Open slot available?
                        If n > 0 Then
                            ' Set item in players inventory
                            Update_InvItemNum Index, n, ItemNum
                            
                            ' Check if it's bind on pickup, if so bind it
                            If Item(ItemNum).Bound = ItemBind.BindOnPickup Then
                                ' Bind it
                                Update_InvItemBound Index, n, True
                            End If
                            
                            ' Check if the item is stackable
                            If Item(ItemNum).Stack Then
                                ' Checks to see if we'll go above max stack for the item
                                If Current_InvItemValue(Index, n) + MapData(MapNum).MapItem(i).Value > Item(ItemNum).StackMax Then
                                    ' Check for more spots in inv
                                    ItemValue = FindNextOpenStack(Index, ItemNum, MapData(MapNum).MapItem(i).Value)
                                   
                                    ' If you don't have enough room in the inv drop it to the ground
                                    ' Else set the rest in your inv
                                    If ItemValue > 0 Then
                                        DeleteItem = False
                                        SendActionMsg MapNum, "You only took " & MapData(MapNum).MapItem(i).Value - ItemValue & " " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
                                    ElseIf ItemValue = 0 Then
                                        DeleteItem = True
                                        SendActionMsg MapNum, "You take " & MapData(MapNum).MapItem(i).Value & " " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
                                    End If
                                Else
                                    Update_InvItemValue Index, n, Current_InvItemValue(Index, n) + MapData(MapNum).MapItem(i).Value
                                    
                                    SendActionMsg MapNum, "You take " & MapData(MapNum).MapItem(i).Value & " " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
                                    DeleteItem = True
                                End If
                            Else
                                Update_InvItemValue Index, n, 1
                                
                                SendActionMsg MapNum, "You take a " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
                                DeleteItem = True
                            End If
                                
                            ' DeleteItem so we can keep the item on the ground
                            If DeleteItem Then
                                ' Erase item from the map
                                MapData(MapNum).MapItem(i).Num = 0
                                MapData(MapNum).MapItem(i).Value = ItemValue
                                MapData(MapNum).MapItem(i).X = 0
                                MapData(MapNum).MapItem(i).Y = 0
                                MapData(MapNum).MapItem(i).DropTime = 0
                                
                                SpawnItemSlot i, 0, 0, MapNum, Current_X(Index), Current_Y(Index), 0
                            Else
                                ' Set the items new value
                                MapData(MapNum).MapItem(i).Value = ItemValue
                                
                                SpawnItemSlot i, MapData(MapNum).MapItem(i).Num, MapData(MapNum).MapItem(i).Value, MapNum, Current_X(Index), Current_Y(Index), MapData(MapNum).MapItem(i).DropTime
                            End If
                            
                            ' Changed to send your whole inventory for stacking
                            SendPlayerInv Index
                            Exit Sub
                        Else
                            SendActionMsg MapNum, "You are fully burdened.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    Next
End Sub

Sub PlayerMapDropItem(ByVal Index As Long, ByVal InvNum As Long, ByVal Amount As Long, ByVal Delete As Boolean)
Dim i As Long
Dim MapNum As Long
Dim ItemNum As Long
Dim ItemVal As Long
Dim DropAmount As Long

    ' Check for subscript out of range
    If Not IsPlaying(Index) Then Exit Sub
    If InvNum <= 0 Then Exit Sub
    If InvNum > MAX_INV Then Exit Sub
    
    ItemNum = Current_InvItemNum(Index, InvNum)
    ItemVal = Current_InvItemValue(Index, InvNum)
    
    If ItemNum <= 0 Then Exit Sub
    If ItemNum > MAX_ITEMS Then Exit Sub
    
    MapNum = Current_Map(Index)
    
    ' Prevent hacking - Check if they are trying to drop more than they had
    If Amount > ItemVal Then Amount = ItemVal
    
    ' set the dropamount
    DropAmount = ItemVal - Amount
    
     ' Check if it's a bound item
    If Current_InvItemBound(Index, InvNum) Then
        If Amount > 1 Then
            SendActionMsg MapNum, "You destroyed " & Amount & " " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
        Else
            SendActionMsg MapNum, "You destroyed " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
        End If
        
        If DropAmount > 0 Then
            Update_InvItem Index, InvNum, ItemNum, DropAmount, True
        Else
            Update_InvItem Index, InvNum, 0, 0, False
        End If
        Exit Sub
    End If
    
    i = FindOpenMapItemSlot(MapNum)
    If i > 0 Then
        MapData(MapNum).MapItem(i).Num = ItemNum
        MapData(MapNum).MapItem(i).X = Current_X(Index)
        MapData(MapNum).MapItem(i).Y = Current_Y(Index)
        MapData(MapNum).MapItem(i).Value = Amount
        
        If Amount > 1 Then
            SendActionMsg MapNum, "You dropped " & Amount & " " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
        Else
            SendActionMsg MapNum, "You dropped a " & Trim$(Item(ItemNum).Name) & ".", ActionColor, ACTIONMSG_SCREEN, 0, 0, Index
        End If
        
        If DropAmount > 0 Then
            Update_InvItem Index, InvNum, ItemNum, DropAmount, False
        Else
            Update_InvItem Index, InvNum, 0, 0, False
        End If
        
        ' Spawn the item before we set the num or we'll get a different free map item slot
        ' Check if you want the item to delete on the map or not
        If Delete Then
            SpawnItemSlot i, MapData(MapNum).MapItem(i).Num, Amount, MapNum, Current_X(Index), Current_Y(Index), GetTickCount
        Else
            SpawnItemSlot i, MapData(MapNum).MapItem(i).Num, Amount, MapNum, Current_X(Index), Current_Y(Index), 0
        End If
    Else
        SendActionMsg MapNum, "No more items can be placed on the ground at this time.", AlertColor, ACTIONMSG_SCREEN, 0, 0, Index
    End If
End Sub

