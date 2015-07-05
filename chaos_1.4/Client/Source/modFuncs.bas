Attribute VB_Name = "modFuncs"
Option Explicit

Function HasItem(itemnum As Long, itemvalue As Long) As Boolean
Dim PlayerHas As Boolean
Dim i

For i = 1 To MAX_INV
If GetPlayerInvItemNum(MyIndex, i) = itemnum Then
If Item(itemnum).Type = ITEM_TYPE_CURRENCY Then
If GetPlayerInvItemValue(MyIndex, i) >= itemvalue Then
PlayerHas = True
Exit For
Else
PlayerHas = False
Exit For
End If
Else
PlayerHas = True
Exit For
End If
End If
Next i

If PlayerHas = True Then
HasItem = True
ElseIf PlayerHas = False Then
HasItem = False
End If

End Function
