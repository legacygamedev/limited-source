Attribute VB_Name = "modPets"
Option Explicit

'Function GetFreePetID()
'Dim i As Long
'    For i = 1 To UBound(Pets) Step 1
'        If Pets(i).owner = 0 Then
'            GetFreePetID = i
'            Exit Function
'        End If
'    Next i
'    GetFreePetID = 0
'End Function
'
'Function loadPet(ByVal index As Long, ByVal PetId As Long, ByVal PetName As String, ByVal PetSprite As Long) As Long
'    Pets(PetId).owner = index
'
'    With Pets(PetId)
'        .Dir = player(.owner).Char(player(.owner).CharNum).Dir
'        .x = player(.owner).Char(player(.owner).CharNum).x
'        .y = player(.owner).Char(player(.owner).CharNum).y
'        .map = player(.owner).Char(player(.owner).CharNum).map
'        .sprite = PetSprite
'        .Name = PetName
'        .target = 0
'    End With
'    loadPet = PetId
'End Function

'
'Sub movePets()
'Dim i As Long
'Dim x As Long
'Dim y As Long
'Dim target As Long
'Dim rndNo As Long
'Dim didwalk As Boolean
'y = 0
'For x = 1 To MAX_PETS Step 1
'    i = x
'    If Pets(i).owner <> 0 Then
'
'
'    ' make sure the pet is on the same map as the player
'    If Pets(i).map = player(Pets(i).owner).Char(player(Pets(i).owner).CharNum).map Then
'
'
'        target = Pets(i).owner
'        rndNo = Int(Rnd * 6)
'        y = player(Pets(i).owner).Char(player(Pets(i).owner).CharNum).map
'        Select Case rndNo
'            Case 0
'                'Up
'                If Pets(i).y > GetPlayerY(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_UP) Then
'                        Call PetMove(y, x, DIR_UP, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Down
'                If Pets(i).y < GetPlayerY(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_DOWN) Then
'                        Call PetMove(y, x, DIR_DOWN, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Left
'                If Pets(i).x > GetPlayerX(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_LEFT) Then
'                        Call PetMove(y, x, DIR_LEFT, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Right
'                If Pets(i).x < GetPlayerX(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_RIGHT) Then
'                        Call PetMove(y, x, DIR_RIGHT, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'            Case 3
'                ' Down
'                If Pets(i).y < GetPlayerY(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_DOWN) Then
'                        Call PetMove(y, x, DIR_DOWN, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Left
'                If Pets(i).x > GetPlayerX(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_LEFT) Then
'                        Call PetMove(y, x, DIR_LEFT, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Right
'                If Pets(i).x < GetPlayerX(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_RIGHT) Then
'                        Call PetMove(y, x, DIR_RIGHT, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                'Up
'                If Pets(i).y > GetPlayerY(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_UP) Then
'                        Call PetMove(y, x, DIR_UP, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'            Case 1
'                ' Left
'                If Pets(i).x > GetPlayerX(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_LEFT) Then
'                        Call PetMove(y, x, DIR_LEFT, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Right
'                If Pets(i).x < GetPlayerX(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_RIGHT) Then
'                        Call PetMove(y, x, DIR_RIGHT, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                'Up
'                If Pets(i).y > GetPlayerY(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_UP) Then
'                        Call PetMove(y, x, DIR_UP, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Down
'                If Pets(i).y < GetPlayerY(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_DOWN) Then
'                        Call PetMove(y, x, DIR_DOWN, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'            Case 2
'                ' Right
'                If Pets(i).x < GetPlayerX(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_RIGHT) Then
'                        Call PetMove(y, x, DIR_RIGHT, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                'Up
'                If Pets(i).y > GetPlayerY(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_UP) Then
'                        Call PetMove(y, x, DIR_UP, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Down
'                If Pets(i).y < GetPlayerY(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_DOWN) Then
'                        Call PetMove(y, x, DIR_DOWN, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'                ' Left
'                If Pets(i).x > GetPlayerX(target) And didwalk = False Then
'                    If CanPetMove(y, x, DIR_LEFT) Then
'                        Call PetMove(y, x, DIR_LEFT, MOVING_WALKING)
'                        didwalk = True
'                    End If
'                End If
'        End Select
'
''        If Not didwalk Then
''            i = Int(Rnd * 2)
''            If i = 1 Then
''                i = Int(Rnd * 4)
''                If CanPetMove(y, x, i) Then
''                    Call PetMove(y, x, i, MOVING_WALKING)
''                End If
''            End If
''
''        End If
'    Else
'        'if its not then warp the pet to the player.
'        Pets(i).x = GetPlayerX(Pets(i).owner)
'        Pets(i).y = GetPlayerY(Pets(i).owner)
'        Pets(i).map = GetPlayerMap(Pets(i).owner)
'        SendUpdatePetToAll (i)
'        DoEvents
'    End If
'    End If
'Next x
'End Sub
'
'
'
'
'Function CanPetMove(ByVal mapNum As Long, ByVal Petnum As Long, ByVal Dir) As Boolean
'Dim i As Long, n As Long
'Dim x As Long, y As Long
'
'    CanPetMove = False
'
'    ' Check for subscript out of range
'    If mapNum <= 0 Or mapNum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Then
'        Exit Function
'    End If
'
'    x = Pets(Petnum).x
'    y = Pets(Petnum).y
'
'    CanPetMove = True
'
'    Select Case Dir
'        Case DIR_UP
'            ' Check to make sure not outside of boundries
'            If y > 0 Then
'                n = map(mapNum).Tile(x, y - 1).type
'
'                ' Check to make sure that the tile is walkable
'                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
'                    CanPetMove = False
'                    Exit Function
'                End If
'
'                ' Check to make sure that there is not a player in the way
'                For i = 1 To MAX_PLAYERS
'                    If IsPlaying(i) Then
'
'                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = x) And (GetPlayerY(i) = (y - 1)) Then
'                            CanPetMove = False
'                            Exit Function
'                        End If
'                    End If
'                Next i
'
'                ' Check to make sure that there is not another npc in the way
'            Else
'                CanPetMove = False
'            End If
'
'        Case DIR_DOWN
'            ' Check to make sure not outside of boundries
'            If y < MAX_MAPY Then
'                n = map(mapNum).Tile(x, y + 1).type
'
'                ' Check to make sure that the tile is walkable
'                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
'                    CanPetMove = False
'                    Exit Function
'                End If
'
'                ' Check to make sure that there is not a player in the way
'                For i = 1 To MAX_PLAYERS
'                    If IsPlaying(i) Then
'                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = x) And (y + 1) Then
'                            CanPetMove = False
'                            Exit Function
'                        End If
'                    End If
'                Next i
'
'                ' Check to make sure that there is not another npc in the way
'            Else
'                CanPetMove = False
'            End If
'
'        Case DIR_LEFT
'            ' Check to make sure not outside of boundries
'            If x > 0 Then
'                n = map(mapNum).Tile(x - 1, y).type
'
'                ' Check to make sure that the tile is walkable
'                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
'                    CanPetMove = False
'                    Exit Function
'                End If
'
'                ' Check to make sure that there is not a player in the way
'                For i = 1 To MAX_PLAYERS
'                    If IsPlaying(i) Then
'                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = x - 1) And (GetPlayerY(i) = y) Then
'                            CanPetMove = False
'                            Exit Function
'                        End If
'                    End If
'                Next i
'
'                ' Check to make sure that there is not another npc in the way
'            Else
'                CanPetMove = False
'            End If
'
'        Case DIR_RIGHT
'            ' Check to make sure not outside of boundries
'            If x < MAX_MAPX Then
'                n = map(mapNum).Tile(x + 1, y).type
'
'                ' Check to make sure that the tile is walkable
'                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM Then
'                    CanPetMove = False
'                    Exit Function
'                End If
'
'                ' Check to make sure that there is not a player in the way
'                For i = 1 To MAX_PLAYERS
'                    If IsPlaying(i) Then
'                        If (GetPlayerMap(i) = mapNum) And (GetPlayerX(i) = x + 1) And (GetPlayerY(i) = y) Then
'                            CanPetMove = False
'                            Exit Function
'                        End If
'                    End If
'                Next i
'
'                ' Check to make sure that there is not another npc in the way
'            Else
'                CanPetMove = False
'            End If
'    End Select
'End Function
'
'Sub PetMove(ByVal mapNum As Long, ByVal Petnum As Long, ByVal Dir As Long, ByVal Movement As Long)
'    Debug.Print "PetMove :" & Dir
'Dim packet As String
'Dim x As Long
'Dim y As Long
'Dim i As Long
'    ' Check for subscript out of range
'    If mapNum <= 0 Or mapNum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then
'        Exit Sub
'    End If
'
'    Pets(Petnum).Dir = Dir
'    'MapNpc(mapNum, mapnpcnum).Dir = Dir
'
'    Select Case Dir
'        Case DIR_UP
'            Pets(Petnum).y = Pets(Petnum).y - 1
'            packet = "PETMOVE" & SEP_CHAR & Petnum & SEP_CHAR & Pets(Petnum).x & SEP_CHAR & Pets(Petnum).y & SEP_CHAR & Pets(Petnum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
'            Call SendDataToMap(mapNum, packet)
'
'        Case DIR_DOWN
'            Pets(Petnum).y = Pets(Petnum).y + 1
'            packet = "PETMOVE" & SEP_CHAR & Petnum & SEP_CHAR & Pets(Petnum).x & SEP_CHAR & Pets(Petnum).y & SEP_CHAR & Pets(Petnum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
'            Call SendDataToMap(mapNum, packet)
'
'        Case DIR_LEFT
'            Pets(Petnum).x = Pets(Petnum).x - 1
'            packet = "PETMOVE" & SEP_CHAR & Petnum & SEP_CHAR & Pets(Petnum).x & SEP_CHAR & Pets(Petnum).y & SEP_CHAR & Pets(Petnum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
'            Call SendDataToMap(mapNum, packet)
'
'        Case DIR_RIGHT
'            Pets(Petnum).x = Pets(Petnum).x + 1
'            packet = "PETMOVE" & SEP_CHAR & Petnum & SEP_CHAR & Pets(Petnum).x & SEP_CHAR & Pets(Petnum).y & SEP_CHAR & Pets(Petnum).Dir & SEP_CHAR & Movement & SEP_CHAR & END_CHAR
'            Call SendDataToMap(mapNum, packet)
'    End Select
'End Sub
