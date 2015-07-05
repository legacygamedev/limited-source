Attribute VB_Name = "modGameLogic"
Option Explicit

Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Const SRCAND = &H8800C6
Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086

Public Const VK_UP = &H26
Public Const VK_DOWN = &H28
Public Const VK_LEFT = &H25
Public Const VK_RIGHT = &H27
Public Const VK_SHIFT = &H10
Public Const VK_RETURN = &HD
Public Const VK_CONTROL = &H11

Public KUp As Long
Public KDown As Long
Public KLeft As Long
Public KRight As Long
Public KAttack As Long
Public KRun As Long
Public KEnter As Long

Public JUp As Long
Public JDown As Long
Public JLeft As Long
Public JRight As Long
Public JAttack As Long
Public JRun As Long
Public JEnter As Long
Public JUpC As Long
Public JDownC As Long
Public JLeftC As Long
Public JRightC As Long
Public JAttackC As Long
Public JRunC As Long
Public JEnterC As Long

Public ID As Boolean

' Menu states
Public Const MENU_STATE_NEWACCOUNT = 0
Public Const MENU_STATE_DELACCOUNT = 1
Public Const MENU_STATE_LOGIN = 2
Public Const MENU_STATE_GETCHARS = 3
Public Const MENU_STATE_NEWCHAR = 4
Public Const MENU_STATE_ADDCHAR = 5
Public Const MENU_STATE_DELCHAR = 6
Public Const MENU_STATE_USECHAR = 7
Public Const MENU_STATE_INIT = 8

' Speed moving vars
Public Const WALK_SPEED = 4
Public Const RUN_SPEED = 8
Public Const GM_WALK_SPEED = 4
Public Const GM_RUN_SPEED = 8
'Set the variable to your desire,
'32 is a safe and recommended setting

' Game direction vars
Public DirUp As Boolean
Public DirDown As Boolean
Public DirLeft As Boolean
Public DirRight As Boolean
Public ShiftDown As Boolean
Public ControlDown As Boolean

' Game text buffer
Public MyText As String

' Index of actual player
Public MyIndex As Long

' Map animation #, used to keep track of what map animation is currently on
Public MapAnim As Byte
Public MapAnimTimer As Long

' Used to freeze controls when getting a new map
Public GettingMap As Boolean

' Used to check if in editor or not and variables for use in editor
Public InEditor As Boolean
Public EditorTileX As Long
Public EditorTileY As Long
Public EditorWarpMap As Long
Public EditorWarpX As Long
Public EditorWarpY As Long

' Used for map item editor
Public ItemEditorNum As Long
Public ItemEditorValue As Long

' Used for map key editor
Public KeyEditorNum As Long
Public KeyEditorTake As Long

' Used for map key opene ditor
Public KeyOpenEditorX As Long
Public KeyOpenEditorY As Long
Public KeyOpenEditorMsg As String

' Map for local use
Public SaveMap As MapRec
Public SaveMapItem() As MapItemRec
Public SaveMapNpc(1 To MAX_MAP_NPCS) As MapNpcRec

' Used for index based editors
Public InItemsEditor As Boolean
Public InNpcEditor As Boolean
Public InShopEditor As Boolean
Public InSpellEditor As Boolean
Public InEmoticonEditor As Boolean
Public EditorIndex As Long

' Game fps
Public GameFPS As Long

' Scrolling Variables
Public NewPlayerX As Long
Public NewPlayerY As Long
Public NewXOffset As Long
Public NewYOffset As Long
Public NewX As Long
Public NewY As Long

' Damage Variables
Public DmgDamage As Long
Public DmgTime As Long
Public NPCDmgDamage As Long
Public NPCDmgTime As Long
Public NPCWho As Long

Public EditorItemX As Long
Public EditorItemY As Long

Public EditorShopNum As Long

Public EditorItemNum1 As Long
Public EditorItemNum2 As Long
Public EditorItemNum3 As Long

Public Arena1 As Long
Public Arena2 As Long
Public Arena3 As Long

Public ii As Long, iii As Long
Public sx As Long

Public MouseDownX As Long
Public MouseDownY As Long

Public SpritePic As Long
Public SpriteItem As Long
Public SpritePrice As Long

Public SoundFileName As String

Public ScreenMode As Byte

Public SignLine1 As String
Public SignLine2 As String
Public SignLine3 As String

Public ClassChange As Long
Public ClassChangeReq As Long

Public NoticeTitle As String
Public NoticeText As String
Public NoticeSound As String

Public ScriptNum As Long

Public Connucted As Boolean

Public InventoryRefresh As Long

Public GameFontSize As Long

Public NCPic As Long

Public SexBlockNum As Long

Public LevelBlockLow As Long
Public LevelBlockHigh As Long
                    
Sub Main()
Dim i As Long
    If (InStr(LCase(Command$), "--startedbypatcher") <= 0) Then End
    ScreenMode = 0
    
    NCPic = 266
    
    Call SetStatus("")
    frmMainMenu.Visible = True
        
    ' Check if the maps directory is there, if its not make it
    If LCase(Dir(App.Path & "\Maps", vbDirectory)) <> "maps" Then
        Call MkDir(App.Path & "\Maps")
    End If
    If UCase(Dir(App.Path & "\GFX", vbDirectory)) <> "GFX" Then
        Call MkDir(App.Path & "\GFX")
    End If
    If UCase(Dir(App.Path & "\Music", vbDirectory)) <> "MUSIC" Then
        Call MkDir(App.Path & "\Music")
    End If
    If UCase(Dir(App.Path & "\SFX", vbDirectory)) <> "SFX" Then
        Call MkDir(App.Path & "\SFX")
    End If
        
    ' Make sure we set that we aren't in the game
    InGame = False
    GettingMap = True
    InEditor = False
    InItemsEditor = False
    InNpcEditor = False
    InShopEditor = False
    InEmoticonEditor = False
    DebugMode = False
    
    frmMirage.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    frmSpriteChange.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
    frmMainMenu.picQuit.Enabled = True
    frmMainMenu.picNewAccount.Enabled = True
    frmMainMenu.picLogin.Enabled = True
    Call TcpInit
End Sub

Sub SetStatus(ByVal Caption As String)
    StatusTime = GetTickCount
    frmMainMenu.lblStatus.Caption = Caption
End Sub

Sub MenuState(ByVal State As Long)
    Connucted = True
    Call SetStatus("Connecting to server...")
    Select Case State
        Case MENU_STATE_NEWACCOUNT
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending new account information...")
                Call SendNewAccount(frmMainMenu.txtNewName.Text, frmMainMenu.txtNewPass.Text)
            End If
                    
        Case MENU_STATE_LOGIN
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending login information...")
                Call SendLogin(frmMainMenu.txtName.Text, frmMainMenu.txtPassword.Text)
            End If
        
        Case MENU_STATE_NEWCHAR
            Call SetStatus("Connected, getting available classes...")
            Call SendGetClasses

        Case MENU_STATE_ADDCHAR
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character addition data...")
                If frmMainMenu.optMale.Value = True Then
                    Call SendAddChar(frmMainMenu.txtCharName, 0, 0, frmMainMenu.lstChars.ListIndex + 1, NCPic)
                Else
                    Call SendAddChar(frmMainMenu.txtCharName, 1, 0, frmMainMenu.lstChars.ListIndex + 1, NCPic)
                End If
            End If
        
        Case MENU_STATE_DELCHAR
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending character deletion request...")
                Call SendDelChar(frmMainMenu.lstChars.ListIndex + 1)
            End If
            
        Case MENU_STATE_USECHAR
            If ConnectToServer = True Then
                Call SetStatus("Connected, sending char data...")
                Call SendUseChar(frmMainMenu.lstChars.ListIndex + 1)
            End If
    End Select

    If Not IsConnected And Connucted = True Then
        Call SetStatus("The server is down!")
    End If
End Sub
Sub GameInit()
    frmMirage.Show
    Call InitDirectX
End Sub

Sub GameDestroy()
    Call DestroyDirectX
    Call StopMidi
    End
End Sub

Sub BltTile(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Side As Long)
Dim Ground As Long
Dim Anim1 As Long
Dim Anim2 As Long
Dim Mask2 As Long
Dim M2Anim As Long
    If MapNum = 0 Then Exit Sub

    If Side = 0 Then
        If x < 22 Then Exit Sub
        If y < 24 Then Exit Sub
        'TopLeft
        With rec_pos
            .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 1 Then
        If y < 24 Then Exit Sub
        'Top
        With rec_pos
            .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = (PIC_X * (MAX_MAPX + 1)) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 2 Then
        If x > 8 Then Exit Sub
        If y < 24 Then Exit Sub
        'TopRight
        With rec_pos
            .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((PIC_X * (MAX_MAPX + 1)) * 2) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 3 Then
        If x < 22 Then Exit Sub
        'Left
        With rec_pos
            .Top = (PIC_Y * (MAX_MAPY + 1)) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 4 Then
        If x < GetPlayerX(MyIndex) - 16 Then Exit Sub
        If x > GetPlayerX(MyIndex) + 16 Then Exit Sub
        If y < GetPlayerY(MyIndex) - 12 Then Exit Sub
        If y > GetPlayerY(MyIndex) + 12 Then Exit Sub
        'Middle
        With rec_pos
            .Top = (PIC_Y * (MAX_MAPY + 1)) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = (PIC_X * (MAX_MAPX + 1)) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 5 Then
        If x > 8 Then Exit Sub
        'Right
        With rec_pos
            .Top = (PIC_Y * (MAX_MAPY + 1)) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((PIC_X * (MAX_MAPX + 1)) * 2) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 6 Then
        If x < 22 Then Exit Sub
        If y > 6 Then Exit Sub
        'BottomLeft
        With rec_pos
            .Top = ((PIC_Y * (MAX_MAPY + 1)) * 2) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 7 Then
        If y > 6 Then Exit Sub
        'Bottom
        With rec_pos
            .Top = ((PIC_Y * (MAX_MAPY + 1)) * 2) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = (PIC_X * (MAX_MAPX + 1)) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 8 Then
        If x > 8 Then Exit Sub
        If y > 6 Then Exit Sub
        'BottomRight
        With rec_pos
            .Top = ((PIC_Y * (MAX_MAPY + 1)) * 2) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((PIC_X * (MAX_MAPX + 1)) * 2) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    End If

    Ground = CheckMap(MapNum).Tile(x, y).Ground
    Anim1 = CheckMap(MapNum).Tile(x, y).Mask
    Anim2 = CheckMap(MapNum).Tile(x, y).Anim
    Mask2 = CheckMap(MapNum).Tile(x, y).Mask2
    M2Anim = CheckMap(MapNum).Tile(x, y).M2Anim
    
    rec.Top = Int(Ground / 14) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Ground - Int(Ground / 14) * 14) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT)
    'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT)
    
    If (MapAnim = 0) Or (Anim2 <= 0) Then
        ' Is there an animation tile to plot?
        If Anim1 > 0 And TempTile(x, y).DoorOpen = NO Then
            rec.Top = Int(Anim1 / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim1 - Int(Anim1 / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If Anim2 > 0 Then
            rec.Top = Int(Anim2 / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Anim2 - Int(Anim2 / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
    
    If (MapAnim = 0) Or (M2Anim <= 0) Then
        ' Is there an animation tile to plot?
        If Mask2 > 0 Then
            rec.Top = Int(Mask2 / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (Mask2 - Int(Mask2 / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    Else
        ' Is there a second animation tile to plot?
        If M2Anim > 0 Then
            rec.Top = Int(M2Anim / 14) * PIC_Y
            rec.Bottom = rec.Top + PIC_Y
            rec.Left = (M2Anim - Int(M2Anim / 14) * 14) * PIC_X
            rec.Right = rec.Left + PIC_X
            Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
            'Call DD_BackBuffer.BltFast((X - NewPlayerX) * PIC_X - NewXOffset, (Y - NewPlayerY) * PIC_Y - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub BltItem(ByVal ItemNum As Long)
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapItem(ItemNum).y * PIC_Y
        .Bottom = .Top + PIC_Y
        .Left = MapItem(ItemNum).x * PIC_X
        .Right = .Left + PIC_X
    End With
    
    rec.Top = Int(Item(MapItem(ItemNum).Num).Pic / 6) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Item(MapItem(ItemNum).Num).Pic - Int(Item(MapItem(ItemNum).Num).Pic / 6) * 6) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    'Call DD_BackBuffer.Blt(rec_pos, DD_ItemSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast((PIC_X * (MAX_MAPX + 1)) + ((MapItem(ItemNum).x - NewPlayerX) * PIC_X + sx - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + ((MapItem(ItemNum).y - NewPlayerY) * PIC_Y + sx - NewYOffset), DD_ItemSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltFringeTile(ByVal MapNum As Long, ByVal x As Long, ByVal y As Long, ByVal Side As Long)
Dim Fringe As Long
Dim FAnim As Long
Dim Fringe2 As Long
Dim F2Anim As Long

    If Side = 0 Then
        If x < 22 Then Exit Sub
        If y < 24 Then Exit Sub
        'TopLeft
        With rec_pos
            .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 1 Then
        If y < 24 Then Exit Sub
        'Top
        With rec_pos
            .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = (PIC_X * (MAX_MAPX + 1)) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 2 Then
        If x > 8 Then Exit Sub
        If y < 24 Then Exit Sub
        'TopRight
        With rec_pos
            .Top = (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((PIC_X * (MAX_MAPX + 1)) * 2) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 3 Then
        If x < 22 Then Exit Sub
        'Left
        With rec_pos
            .Top = (PIC_Y * (MAX_MAPY + 1)) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 4 Then
        'Middle
        With rec_pos
            .Top = (PIC_Y * (MAX_MAPY + 1)) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = (PIC_X * (MAX_MAPX + 1)) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 5 Then
        If x > 8 Then Exit Sub
        'Right
        With rec_pos
            .Top = (PIC_Y * (MAX_MAPY + 1)) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((PIC_X * (MAX_MAPX + 1)) * 2) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 6 Then
        If x < 22 Then Exit Sub
        If y > 6 Then Exit Sub
        'BottomLeft
        With rec_pos
            .Top = ((PIC_Y * (MAX_MAPY + 1)) * 2) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 7 Then
        If y > 6 Then Exit Sub
        'Bottom
        With rec_pos
            .Top = ((PIC_Y * (MAX_MAPY + 1)) * 2) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = (PIC_X * (MAX_MAPX + 1)) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    ElseIf Side = 8 Then
        If x > 8 Then Exit Sub
        If y > 6 Then Exit Sub
        'BottomRight
        With rec_pos
            .Top = ((PIC_Y * (MAX_MAPY + 1)) * 2) + (y - NewPlayerY) * PIC_Y + sx - NewYOffset
            .Bottom = .Top + PIC_Y
            .Left = ((PIC_X * (MAX_MAPX + 1)) * 2) + ((x - NewPlayerX) * PIC_X + sx - NewXOffset)
            .Right = .Left + PIC_X
        End With
    End If
    
    Fringe = CheckMap(MapNum).Tile(x, y).Fringe
    FAnim = CheckMap(MapNum).Tile(x, y).FAnim
    Fringe2 = CheckMap(MapNum).Tile(x, y).Fringe2
    F2Anim = CheckMap(MapNum).Tile(x, y).F2Anim
        
    If (MapAnim = 0) Or (FAnim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe > 0 Then
        rec.Top = Int(Fringe / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (Fringe - Int(Fringe / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        'Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
    
        If FAnim > 0 Then
        rec.Top = Int(FAnim / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (FAnim - Int(FAnim / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        'Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If

    If (MapAnim = 0) Or (F2Anim <= 0) Then
        ' Is there an animation tile to plot?
        
        If Fringe2 > 0 Then
        rec.Top = Int(Fringe2 / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (Fringe2 - Int(Fringe2 / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        'Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    
    Else
    
        If F2Anim > 0 Then
        rec.Top = Int(F2Anim / 14) * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (F2Anim - Int(F2Anim / 14) * 14) * PIC_X
        rec.Right = rec.Left + PIC_X
        Call DD_BackBuffer.Blt(rec_pos, DD_TileSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        'Call DD_BackBuffer.BltFast((x - NewPlayerX) * PIC_X + sx - NewXOffset, (y - NewPlayerY) * PIC_Y + sx - NewYOffset, DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
        
    End If
End Sub

Sub BltPlayer(ByVal Index As Long)
Dim Anim As Byte
Dim x As Long, y As Long

If Index = MyIndex Then
   
    ' Check for animation
    Anim = 0
    If Player(MyIndex).Attacking = 0 Then
        Select Case GetPlayerDir(MyIndex)
            Case DIR_UP
                If (Player(MyIndex).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(MyIndex).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(MyIndex).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(MyIndex).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(MyIndex).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(MyIndex).AttackTimer + 1000 < GetTickCount Then
        Player(MyIndex).Attacking = 0
        Player(MyIndex).AttackTimer = 0
    End If
    
    rec.Top = GetPlayerSprite(MyIndex) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (GetPlayerDir(MyIndex) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    x = (PIC_X * (MAX_MAPX + 1)) + NewX + sx
    y = (PIC_Y * (MAX_MAPY + 1)) + NewY + sx
        
    Call DD_BackBuffer.BltFast(x, y, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
Else
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset
        .Bottom = .Top + PIC_Y
        .Left = GetPlayerX(Index) * PIC_X + Player(Index).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If Player(Index).Attacking = 0 Then
        Select Case GetPlayerDir(Index)
            Case DIR_UP
                If (Player(Index).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (Player(Index).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (Player(Index).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (Player(Index).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If Player(Index).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If Player(Index).AttackTimer + 1000 < GetTickCount Then
        Player(Index).Attacking = 0
        Player(Index).AttackTimer = 0
    End If
    
    rec.Top = GetPlayerSprite(Index) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (GetPlayerDir(Index) * 3 + Anim) * PIC_X
    rec.Right = rec.Left + PIC_X
    
    x = (PIC_X * (MAX_MAPX + 1)) + (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset)
    y = (PIC_Y * (MAX_MAPY + 1)) + (GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset) '- 4
    
    ' Check if its out of bounds because of the offset
    If y < 0 Then
        y = 0
        rec.Top = rec.Top + (y * -1)
    End If
        
    'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
    Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End If
End Sub

Sub BltNpc(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If
        
    If Npc(MapNpc(MapNpcNum).Num).Big = 0 Then
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * PIC_X
        rec.Right = rec.Left + PIC_X
        
        x = (PIC_X * (MAX_MAPX + 1)) + (MapNpc(MapNpcNum).x * PIC_X + sx + MapNpc(MapNpcNum).XOffset)
        y = (PIC_Y * (MAX_MAPY + 1)) + (MapNpc(MapNpcNum).y * PIC_Y + sx + MapNpc(MapNpcNum).YOffset)
        
        ' Check if its out of bounds because of the offset
        If y < 0 Then
            y = 0
            rec.Top = rec.Top + (y * -1)
        End If
            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x - (NewPlayerX * PIC_X) - NewXOffset, y - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    Else
        rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
        rec.Bottom = rec.Top + 32
        rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
        rec.Right = rec.Left + 64
    
        x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
        y = MapNpc(MapNpcNum).y * 32 + sx + MapNpc(MapNpcNum).YOffset
   
        If y < 0 Then
            rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
            rec.Bottom = rec.Top + 32
            y = MapNpc(MapNpcNum).YOffset + sx
        End If
        
        If x < 0 Then
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
            rec.Right = rec.Left + 48
            x = MapNpc(MapNpcNum).XOffset + sx
        End If
        
        If x > MAX_MAPX * 32 Then
            rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
            rec.Right = rec.Left + 48
            x = MAX_MAPX * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
        End If

        Call DD_BackBuffer.BltFast((PIC_X * (MAX_MAPX + 1)) + (x - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (y - (NewPlayerY * PIC_Y) - NewYOffset), DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltNpcTop(ByVal MapNpcNum As Long)
Dim Anim As Byte
Dim x As Long, y As Long

    ' Make sure that theres an npc there, and if not exit the sub
    If MapNpc(MapNpcNum).Num <= 0 Then
        Exit Sub
    End If
    
    If Npc(MapNpc(MapNpcNum).Num).Big = 0 Then Exit Sub
    
    ' Only used if ever want to switch to blt rather then bltfast
    With rec_pos
        .Top = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).YOffset
        .Bottom = .Top + PIC_Y
        .Left = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).XOffset
        .Right = .Left + PIC_X
    End With
    
    ' Check for animation
    Anim = 0
    If MapNpc(MapNpcNum).Attacking = 0 Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_DOWN
                If (MapNpc(MapNpcNum).YOffset < PIC_Y / 2 * -1) Then Anim = 1
            Case DIR_LEFT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2) Then Anim = 1
            Case DIR_RIGHT
                If (MapNpc(MapNpcNum).XOffset < PIC_Y / 2 * -1) Then Anim = 1
        End Select
    Else
        If MapNpc(MapNpcNum).AttackTimer + 500 > GetTickCount Then
            Anim = 2
        End If
    End If
    
    ' Check to see if we want to stop making him attack
    If MapNpc(MapNpcNum).AttackTimer + 1000 < GetTickCount Then
        MapNpc(MapNpcNum).Attacking = 0
        MapNpc(MapNpcNum).AttackTimer = 0
    End If
    
    rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * PIC_Y
        
     rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64
     rec.Bottom = rec.Top + 32
     rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
     rec.Right = rec.Left + 64
 
     x = MapNpc(MapNpcNum).x * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
     y = MapNpc(MapNpcNum).y * 32 + sx - 32 + MapNpc(MapNpcNum).YOffset

     If y < 0 Then
         rec.Top = Npc(MapNpc(MapNpcNum).Num).Sprite * 64 + 32
         rec.Bottom = rec.Top
         y = MapNpc(MapNpcNum).YOffset + sx
     End If
     
     If x < 0 Then
         rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64 + 16
         rec.Right = rec.Left + 48
         x = MapNpc(MapNpcNum).XOffset + sx
     End If
     
     If x > MAX_MAPX * 32 Then
         rec.Left = (MapNpc(MapNpcNum).Dir * 3 + Anim) * 64
         rec.Right = rec.Left + 48
         x = MAX_MAPX * 32 + sx - 16 + MapNpc(MapNpcNum).XOffset
     End If

     Call DD_BackBuffer.BltFast((PIC_X * (MAX_MAPX + 1)) + (x - (NewPlayerX * PIC_X) - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (y - (NewPlayerY * PIC_Y) - NewYOffset), DD_BigSpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long
    
    ' Check access level
    If GetPlayerPK(Index) = NO Then
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(White)
            Case Is > 1
                Color = RGB(153, 255, 153)
        End Select
    Else
        Select Case GetPlayerAccess(Index)
            Case 0
                Color = QBColor(BrightRed)
            Case Is > 1
                Color = RGB(153, 255, 153)
        End Select
    End If
  
If Index = MyIndex Then
    TextX = (PIC_X * (MAX_MAPX + 1)) + (NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerName(MyIndex)) / 2) * GameFontSize))
    TextY = (PIC_Y * (MAX_MAPY + 1)) + (NewY + sx - Int(PIC_Y / 2))
    Call DrawText(TexthDC, TextX, TextY, GetPlayerName(MyIndex), Color)
Else
    TextX = (PIC_X * (MAX_MAPX + 1)) + (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerName(Index)) / 2) * GameFontSize))
    TextY = (PIC_Y * (MAX_MAPY + 1)) + (GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2))
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerName(Index), Color)
End If
End Sub

Sub BltPlayerGuildName(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim Color As Long

    ' Check access level
    Select Case GetPlayerGuildAccess(Index)
        Case 0
            If GetPlayerSTR(Index) > 0 Then
                Color = QBColor(Red)
            Else
                Color = QBColor(Red)
            End If
        Case 1
            Color = QBColor(BrightCyan)
        Case 2
            Color = QBColor(Pink)
        Case 3
            Color = QBColor(BrightGreen)
        Case 4
            Color = QBColor(Yellow)
    End Select

If Index = MyIndex Then
    TextX = (PIC_X * (MAX_MAPX + 1)) + (NewX + sx + Int(PIC_X / 2) - ((Len(GetPlayerGuild(MyIndex)) / 2) * GameFontSize))
    TextY = (PIC_Y * (MAX_MAPY + 1)) + (NewY + sx - Int(PIC_Y / 4) - 20)
    
    Call DrawText(TexthDC, TextX, TextY, GetPlayerGuild(MyIndex), Color)
Else
    ' Draw name
    TextX = (PIC_X * (MAX_MAPX + 1)) + (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + Int(PIC_X / 2) - ((Len(GetPlayerGuild(Index)) / 2) * GameFontSize))
    TextY = (PIC_Y * (MAX_MAPY + 1)) + (GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - Int(PIC_Y / 2) - 12)
    Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) - NewXOffset, TextY - (NewPlayerY * PIC_Y) - NewYOffset, GetPlayerGuild(Index), Color)
End If
End Sub

Sub ProcessMovement(ByVal Index As Long)
    ' Check if player is walking, and if so process moving them over
If Player(Index).Moving = MOVING_WALKING Then
        If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - GM_WALK_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + GM_WALK_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - GM_WALK_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + GM_WALK_SPEED
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - WALK_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + WALK_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - WALK_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + WALK_SPEED
            End Select
        End If
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then
            Player(Index).Moving = 0
        End If
    End If

    ' Check if player is running, and if so process moving them over
If Player(Index).Moving = MOVING_RUNNING Then
            If Player(Index).Access > 0 Then
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - GM_RUN_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + GM_RUN_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - GM_RUN_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + GM_RUN_SPEED
            End Select
        Else
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    Player(Index).YOffset = Player(Index).YOffset - RUN_SPEED
                Case DIR_DOWN
                    Player(Index).YOffset = Player(Index).YOffset + RUN_SPEED
                Case DIR_LEFT
                    Player(Index).XOffset = Player(Index).XOffset - RUN_SPEED
                Case DIR_RIGHT
                    Player(Index).XOffset = Player(Index).XOffset + RUN_SPEED
            End Select
        End If
        
        ' Check if completed walking over to the next tile
        If (Player(Index).XOffset = 0) And (Player(Index).YOffset = 0) Then
            Player(Index).Moving = 0
        End If
    End If
End Sub

Sub ProcessNpcMovement(ByVal MapNpcNum As Long)
    ' Check if npc is walking, and if so process moving them over
    If MapNpc(MapNpcNum).Moving = MOVING_WALKING Then
        Select Case MapNpc(MapNpcNum).Dir
            Case DIR_UP
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset - WALK_SPEED
            Case DIR_DOWN
                MapNpc(MapNpcNum).YOffset = MapNpc(MapNpcNum).YOffset + WALK_SPEED
            Case DIR_LEFT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset - WALK_SPEED
            Case DIR_RIGHT
                MapNpc(MapNpcNum).XOffset = MapNpc(MapNpcNum).XOffset + WALK_SPEED
        End Select
        
        ' Check if completed walking over to the next tile
        If (MapNpc(MapNpcNum).XOffset = 0) And (MapNpc(MapNpcNum).YOffset = 0) Then
            MapNpc(MapNpcNum).Moving = 0
        End If
    End If
End Sub

Sub HandleKeypresses(ByVal KeyAscii As Integer)
Dim ChatText As String
Dim name As String
Dim i As Long
Dim n As Long
Dim s As String

    MyText = frmMirage.txtMyTextBox.Text

    ' Handle when the player presses the return key
    If (KeyAscii = vbKeyReturn) Then
        If Player(MyIndex).y - 1 > -1 Then
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN And Player(MyIndex).Dir = DIR_UP Then
                Call AddText("The Sign Reads:", Black)
                If Trim(CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1) <> "" Then
                    Call AddText(Trim(CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String1), Grey)
                End If
                If Trim(CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2) <> "" Then
                    Call AddText(Trim(CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String2), Grey)
                End If
                If Trim(CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3) <> "" Then
                    Call AddText(Trim(CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).String3), Grey)
                End If
            Exit Sub
            End If
        End If

        ' Broadcast message
        If Mid(MyText, 1, 1) = "'" Then
            If frmMirage.chkBroadcast.Value = Checked Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call BroadcastMsg(ChatText)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            Else
                Call AddText("Your turned off broadcast chat!", BrightRed)
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        End If
        
        ' Emote message
        If Mid(MyText, 1, 1) = "-" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call EmoteMsg(ChatText)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Guild message
        If Mid(MyText, 1, 1) = "@" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("guildchat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Guild message
        If Mid(MyText, 1, 6) = "/gchat" Then
            ChatText = Mid(MyText, 7, Len(MyText) - 6)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("guildchat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Party message
        If Mid(MyText, 1, 1) = "#" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("partychat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Party message
        If Mid(MyText, 1, 6) = "/pchat" Then
            ChatText = Mid(MyText, 7, Len(MyText) - 6)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("partychat" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Reply message
        If Mid(MyText, 1, 6) = "/reply" Then
            ChatText = Mid(MyText, 7, Len(MyText) - 6)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("replymsg" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If

        ' Player message
        If Mid(MyText, 1, 1) = "!" Then
            ChatText = Mid(MyText, 2, Len(MyText) - 1)
            name = ""
                    
            ' Get the desired player from the user text
            For i = 1 To Len(ChatText)
                If Mid(ChatText, i, 1) <> " " Then
                    name = name & Mid(ChatText, i, 1)
                Else
                    Exit For
                End If
            Next i
                    
            ' Make sure they are actually sending something
            If Len(ChatText) - i > 0 Then
                ChatText = Mid(ChatText, i + 1, Len(ChatText) - i)
                    
                ' Send the message to the player
                Call PlayerMsg(ChatText, name)
            Else
                Call AddText("Usage: !playername msghere", AlertColor)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
            
        ' // Commands //
        If LCase(Mid(MyText, 1, 5)) = "/pass" Then
            If Len(MyText) > 10 Then
                ChatText = Mid(MyText, 6, Len(MyText) - 5)
                Call SendData("changepass" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /pass passwordhere", AlertColor)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Display all emotes
        If LCase(Mid(MyText, 1, 7)) = "/emotes" Then
            For i = 0 To MAX_EMOTICONS
                If Trim(Emoticons(i).Command) <> "/" Then
                    If i = 0 Then
                        s = Emoticons(i).Command
                    Else
                        s = s & ", " & Emoticons(i).Command
                    End If
                End If
            Next i
            Call AddText("Emoticons: " & s, AlertColor)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If

        ' Kick player from guild
        If LCase(Mid(MyText, 1, 10)) = "/guildkick" Then
            If Len(MyText) > 10 Then
                ChatText = Mid(MyText, 11, Len(MyText) - 10)
                Call SendData("kickfromguild" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /guildkick playernamehere", AlertColor)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 9)) = "/guildwho" Then
            Call SendData("guildwho" & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 11)) = "/guildleave" Then
            Call SendData("guildleave" & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Invite player to guild
        If LCase(Mid(MyText, 1, 12)) = "/guildinvite" Then
            If Len(MyText) > 12 Then
                ChatText = Mid(MyText, 13, Len(MyText) - 12)
                Call SendData("invitetoguild" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /guildinvite playernamehere", AlertColor)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 12)) = "/guildaccept" Then
            Call SendData("guildinvite" & SEP_CHAR & 0 & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 13)) = "/guilddecline" Then
            Call SendData("guildinvite" & SEP_CHAR & 1 & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
  
        ' Buy guild
        If LCase(Mid(MyText, 1, 12)) = "/createguild" Then
            If Len(MyText) > 12 Then
                ChatText = Mid(MyText, 13, Len(MyText) - 12)
                Call SendData("buyguild" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            Else
                Call AddText("Usage: /createguild guildnamehere", AlertColor)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Training
        If LCase(Mid(MyText, 1, 6)) = "/train" Then
            Call SendData("traininghouse" & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Call upon admins
        If LCase(Mid(MyText, 1, 11)) = "/calladmins" Then
            Call SendData("calladmins" & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Verification User
        'If LCase(Mid(MyText, 1, 5)) = "/info" Then
        '    ChatText = Mid(MyText, 6, Len(MyText) - 5)
        '    Call SendData("playerinforequest" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
        '    frmMirage.txtMyTextBox.Text = ""
        '    Exit Sub
        'End If
        
        ' Whos Online
        If LCase(Mid(MyText, 1, 4)) = "/who" Then
            Call SendWhosOnline
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
                        
        ' Party invite request
        If LCase(Mid(MyText, 1, 7)) = "/invite" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 8 Then
                ChatText = Mid(MyText, 9, Len(MyText) - 8)
                Call SendPartyInvite(ChatText)
            Else
                Call AddText("Usage: /invite <playernamehere>", AlertColor)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
                        
        ' Show inventory
        If LCase(Mid(MyText, 1, 4)) = "/inv" Then
            frmMirage.picInv3.ZOrder (0)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Request stats
        If LCase(Mid(MyText, 1, 6)) = "/stats" Then
            Call SendData("getstats" & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
         
        ' Refresh Player
        If LCase(Mid(MyText, 1, 8)) = "/refresh" Then
            Call SendData("refresh" & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If

        ' Decline Chat
        If LCase(Mid(MyText, 1, 12)) = "/chatdecline" Then
            Call SendData("dchat" & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Accept Chat
        If LCase(Mid(MyText, 1, 5)) = "/chat" Then
            Call SendData("achat" & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        If LCase(Mid(MyText, 1, 6)) = "/trade" Then
            ' Make sure they are actually sending something
            Call AddText("Player To Player Trade Disabled!", AlertColor)
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                'Call SendTradeRequest(ChatText)
            Else
                'Call AddText("Usage: /trade playernamehere", AlertColor)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If

        ' Accept Trade
        If LCase(Mid(MyText, 1, 7)) = "/accept" Then
            'Call SendAcceptTrade
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Decline Trade
        If LCase(Mid(MyText, 1, 8)) = "/decline" Then
            'Call SendDeclineTrade
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Party request
        If LCase(Mid(MyText, 1, 6)) = "/party" Then
            ' Make sure they are actually sending something
            If Len(MyText) > 7 Then
                ChatText = Mid(MyText, 8, Len(MyText) - 7)
                Call SendPartyRequest(ChatText)
            Else
                Call AddText("Usage: /party playernamehere", AlertColor)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Join party
        If LCase(Mid(MyText, 1, 5)) = "/join" Then
            Call SendJoinParty
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' Leave party
        If LCase(Mid(MyText, 1, 6)) = "/leave" Then
            Call SendLeaveParty
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
        
        ' // Moniter Admin Commands //
        If GetPlayerAccess(MyIndex) > 0 Then
        
             ' Warping to a player
            If LCase(Mid(MyText, 1, 9)) = "/warpmeto" Then
                If Len(MyText) > 10 Then
                    MyText = Mid(MyText, 10, Len(MyText) - 9)
                    Call WarpMeTo(MyText)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
                        
            ' Warping a player to you
            If LCase(Mid(MyText, 1, 9)) = "/warptome" Then
                If Len(MyText) > 10 Then
                    MyText = Mid(MyText, 10, Len(MyText) - 9)
                    Call WarpToMe(MyText)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        
            ' day night command
            If LCase(Mid(MyText, 1, 9)) = "/daynight" Then
                If GameTime = TIME_DAY Then
                    GameTime = TIME_NIGHT
                Else
                    GameTime = TIME_DAY
                End If
                Call SendData("GmTime" & SEP_CHAR & GameTime & SEP_CHAR & END_CHAR)
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            If LCase(Mid(MyText, 1, 10)) = "/intensity" Then
                MyText = Mid(MyText, 11, Len(MyText) - 10)
                If IsNumeric(MyText) = True Then
                    If Val(MyText) <= 0 Or Val(MyText) > 50 Then
                        Call AddText("Please enter a number within 1 and 50!", BrightRed)
                        Exit Sub
                    End If
                    Call SendData("intensity" & SEP_CHAR & Val(MyText) & SEP_CHAR & END_CHAR)
                Else
                    Call AddText("Please use numbers only!", BrightRed)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' weather command
            If LCase(Mid(MyText, 1, 8)) = "/weather" Then
                If Len(MyText) > 8 Then
                    MyText = Mid(MyText, 9, Len(MyText) - 8)
                    If IsNumeric(MyText) = True Then
                        Call SendData("weather" & SEP_CHAR & Val(MyText) & SEP_CHAR & END_CHAR)
                    Else
                        If Trim(LCase(MyText)) = "none" Then i = 0
                        If Trim(LCase(MyText)) = "rain" Then i = 1
                        If Trim(LCase(MyText)) = "snow" Then i = 2
                        If Trim(LCase(MyText)) = "thunder" Then i = 3
                        Call SendData("weather" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    End If
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            If LCase(Mid(MyText, 1, 7)) = "/unjail" Then
                If Len(MyText) > 8 Then
                    ChatText = Mid(MyText, 9, Len(MyText) - 8)
                    Call SendData("unjailplayer" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
                Else
                    Call AddText("Usage: /unjail playernamehere", AlertColor)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            If LCase(Mid(MyText, 1, 5)) = "/jail" Then
                If Len(MyText) > 6 Then
                    ChatText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendData("jailplayer" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
                Else
                    Call AddText("Usage: /jail playernamehere", AlertColor)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        
            'mute player
            If LCase(Mid(MyText, 1, 5)) = "/mute" Then
                ' Make sure they are actually sending something
                If Len(MyText) > 6 Then
                    ChatText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendData("muteplayer" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
                Else
                    Call AddText("Usage: /mute playernamehere", AlertColor)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        
            ' Mute All
            If LCase(Mid(MyText, 1, 9)) = "/massmute" Then
                Call SendData("mutebroadcast" & SEP_CHAR & END_CHAR)
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
                        
            ' Kicking a player
            If LCase(Mid(MyText, 1, 5)) = "/kick" Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    Call SendKick(MyText)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
    
            ' Global Message
            If Mid(MyText, 1, 1) = """" Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call GlobalMsg(ChatText)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        
            ' Admin Message
            If Mid(MyText, 1, 1) = "=" Then
                ChatText = Mid(MyText, 2, Len(MyText) - 1)
                If Len(Trim(ChatText)) > 0 Then
                    Call AdminMsg(ChatText)
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
             ' Check the ban list
            If Mid(MyText, 1, 8) = "/banlist" Then
                Call SendBanList
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Banning a player
            If LCase(Mid(MyText, 1, 4)) = "/ban" Then
                If Len(MyText) > 5 Then
                    MyText = Mid(MyText, 6, Len(MyText) - 5)
                    Call SendBan(MyText)
                    frmMirage.txtMyTextBox.Text = ""
                End If
                Exit Sub
            End If
        End If
        
        ' // Mapper Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_MAPPER Then

            ' Location
            If LCase(Mid(MyText, 1, 4)) = "/loc" Then
                Call SendRequestLocation
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Map Editor
            If LCase(Mid(MyText, 1, 10)) = "/mapeditor" Then
                Call SendRequestEditMap
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Map report
            If LCase(Mid(MyText, 1, 10)) = "/mapreport" Then
                Call SendData("mapreport" & SEP_CHAR & END_CHAR)
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
                        
            ' Warping to a map
            If LCase(Mid(MyText, 1, 7)) = "/warpto" Then
                If Len(MyText) > 8 Then
                    MyText = Mid(MyText, 8, Len(MyText) - 7)
                    n = Val(MyText)
                
                    ' Check to make sure its a valid map #
                    If n > 0 And n <= MAX_MAPS Then
                        Call WarpTo(n)
                    Else
                        Call AddText("Invalid map number.", Red)
                    End If
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Setting sprite
            If LCase(Mid(MyText, 1, 10)) = "/setsprite" Then
                If Len(MyText) > 11 Then
                    ' Get sprite #
                    MyText = Mid(MyText, 12, Len(MyText) - 11)
                
                    Call SendSetSprite(Val(MyText))
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        
            ' Respawn request
            If Mid(MyText, 1, 8) = "/respawn" Then
                Call SendMapRespawn
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        
            ' MOTD change
            If Mid(MyText, 1, 5) = "/motd" Then
                If Len(MyText) > 6 Then
                    MyText = Mid(MyText, 7, Len(MyText) - 6)
                    If Trim(MyText) <> "" Then
                        Call SendMOTDChange(MyText)
                    End If
                End If
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        End If
            
        ' // Developer Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_DEVELOPER Then
            ' Editing item request
            If Mid(MyText, 1, 9) = "/edititem" Then
                Call SendRequestEditItem
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Editing emoticon request
            If Mid(MyText, 1, 13) = "/editemoticon" Then
                Call SendRequestEditEmoticon
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Editing npc request
            If Mid(MyText, 1, 8) = "/editnpc" Then
                Call SendRequestEditNpc
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Editing shop request
            If Mid(MyText, 1, 9) = "/editshop" Then
                Call SendRequestEditShop
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        
            ' Editing spell request
            If LCase(Trim(MyText)) = "/editspell" Then
            'If Mid(MyText, 1, 10) = "/editspell" Then
                Call SendRequestEditSpell
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        End If
        
        ' // Creator Admin Commands //
        If GetPlayerAccess(MyIndex) >= ADMIN_CREATOR Then
            If LCase(Mid(MyText, 1, 9)) = "/masswarp" Then
                Call SendData("masswarp" & SEP_CHAR & END_CHAR)
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Giving another player access
            If LCase(Mid(MyText, 1, 10)) = "/setaccess" Then
                ' Get access #
                i = Val(Mid(MyText, 12, 1))
                
                MyText = Mid(MyText, 14, Len(MyText) - 13)
                
                Call SendSetAccess(MyText, i)
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
            
            ' Ban destroy
            If LCase(Mid(MyText, 1, 15)) = "/destroybanlist" Then
                Call SendBanDestroy
                frmMirage.txtMyTextBox.Text = ""
                Exit Sub
            End If
        End If
        
        ' Reply message
        If Mid(MyText, 1, 2) = "/r" Then
            ChatText = Mid(MyText, 3, Len(MyText) - 2)
            If Len(Trim(ChatText)) > 0 Then
                Call SendData("replymsg" & SEP_CHAR & ChatText & SEP_CHAR & END_CHAR)
            End If
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
 
        ' Tell them its not a valid command
        If Left$(Trim(MyText), 1) = "/" Then
            For i = 0 To MAX_EMOTICONS
                If Trim(Emoticons(i).Command) = Trim(MyText) And Trim(Emoticons(i).Command) <> "/" Then
                    Call SendData("checkemoticons" & SEP_CHAR & i & SEP_CHAR & END_CHAR)
                    frmMirage.txtMyTextBox.Text = ""
                Exit Sub
                End If
            Next i
            Call SendData("checkcommands" & SEP_CHAR & MyText & SEP_CHAR & END_CHAR)
            frmMirage.txtMyTextBox.Text = ""
            Exit Sub
        End If
            
        ' Say message
        If Len(Trim(MyText)) > 0 Then
            Call SayMsg(MyText)
        End If
        frmMirage.txtMyTextBox.Text = ""
        Exit Sub
    End If
    
    ' Handle when the user presses the backspace key
    'If (KeyAscii = vbKeyBack) Then
        'If Len(MyText) > 0 Then
            'MyText = Mid(MyText, 1, Len(MyText) - 1)
        'End If
    'End If
    
    ' And if neither, then add the character to the user's text buffer
    'If (KeyAscii <> vbKeyReturn) And (KeyAscii <> vbKeyBack) Then
        ' Make sure they just use standard keys, no gay shitty macro keys
        'If KeyAscii >= 32 And KeyAscii <= 126 Then
            'MyText = MyText & Chr(KeyAscii)
        'End If
    'End If
End Sub

Sub CheckMapGetItem()
    If GetTickCount > Player(MyIndex).MapGetTimer + 250 And Trim(MyText) = "" Then
        Player(MyIndex).MapGetTimer = GetTickCount
        Call SendData("mapgetitem" & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub CheckAttack()
    If ControlDown = True And Player(MyIndex).AttackTimer + 1000 < GetTickCount And Player(MyIndex).Attacking = 0 Then
        Player(MyIndex).Attacking = 1
        Player(MyIndex).AttackTimer = GetTickCount
        Call SendData("attack" & SEP_CHAR & END_CHAR)
    End If
End Sub

Sub CheckInput(ByVal KeyState As Byte, ByVal KeyCode As Integer, ByVal Shift As Integer)
    If GettingMap = False Then
        If KeyState = 1 Then
            If KeyCode = vbKeyReturn Then
                Call CheckMapGetItem
            End If
            If KeyCode = vbKeyControl Then
                ControlDown = True
            End If
            If KeyCode = vbKeyUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            End If
            If KeyCode = vbKeyLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            End If
            If KeyCode = vbKeyRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            End If
            If KeyCode = vbKeyShift Then
                ShiftDown = True
            End If
        Else
            If KeyCode = vbKeyUp Then DirUp = False
            If KeyCode = vbKeyDown Then DirDown = False
            If KeyCode = vbKeyLeft Then DirLeft = False
            If KeyCode = vbKeyRight Then DirRight = False
            If KeyCode = vbKeyShift Then ShiftDown = False
            If KeyCode = vbKeyControl Then ControlDown = False
        End If
    End If
End Sub

Function IsTryingToMove() As Boolean
    If (DirUp = True) Or (DirDown = True) Or (DirLeft = True) Or (DirRight = True) Then
        IsTryingToMove = True
    Else
        IsTryingToMove = False
    End If
End Function

Function CanMove() As Boolean
Dim i As Long, d As Long

    CanMove = True
    
    ' Make sure they aren't trying to move when they are already moving
    If Player(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' Make sure they haven't just casted a spell
    If Player(MyIndex).CastedSpell = YES Then
        If GetTickCount > Player(MyIndex).AttackTimer + 1000 Then
            Player(MyIndex).CastedSpell = NO
        Else
            CanMove = False
            Exit Function
        End If
    End If
    
    d = GetPlayerDir(MyIndex)
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            ' Check for level block
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_LEVEL_BLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data1 >= Player(MyIndex).Level Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data2 <= Player(MyIndex).Level Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DirUp Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check for sex block
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SEX_BLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data1 = Player(MyIndex).Sex Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DirUp Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If

            ' Check to see if the map tile is blocked or not
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_BLOCKED Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_UP Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
            
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_CBLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data1 = Player(MyIndex).Class Then Exit Function
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data2 = Player(MyIndex).Class Then Exit Function
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
            End If
                                                    
            ' Check to see if the key door is open or not
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_KEY Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) - 1).DoorOpen = NO Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_UP Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) Then
                    If GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                        If (GetPlayerX(i) = GetPlayerX(MyIndex)) And (GetPlayerY(i) = GetPlayerY(MyIndex) - 1) Then
                            CanMove = False
                        
                            ' Set the new direction if they weren't facing that direction
                            If d <> DIR_UP Then
                                Call SendPlayerDir
                            End If
                            Exit Function
                        End If
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex)) And (MapNpc(i).y = GetPlayerY(MyIndex) - 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_UP Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If CheckMap(GetPlayerMap(MyIndex)).Up > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
            
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < MAX_MAPY Then
            ' Check for level block
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_LEVEL_BLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data1 >= Player(MyIndex).Level Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data2 <= Player(MyIndex).Level Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DirDown Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check for sex block
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_SEX_BLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data1 = Player(MyIndex).Sex Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DirDown Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check to see if the map tile is blocked or not
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_BLOCKED Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_CBLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data1 = Player(MyIndex).Class Then Exit Function
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data2 = Player(MyIndex).Class Then Exit Function
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_DOWN Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_KEY Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex), GetPlayerY(MyIndex) + 1).DoorOpen = NO Then
                    CanMove = False
                
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_DOWN Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex)) And (GetPlayerY(i) = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_DOWN Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
            
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex)) And (MapNpc(i).y = GetPlayerY(MyIndex) + 1) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_DOWN Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If CheckMap(GetPlayerMap(MyIndex)).Down > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
                
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            ' Check for level block
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_LEVEL_BLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data1 >= Player(MyIndex).Level Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data2 <= Player(MyIndex).Level Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DirLeft Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check for sex block
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SEX_BLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Sex Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DirLeft Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check to see if the map tile is blocked or not
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_CBLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Class Then Exit Function
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data2 = Player(MyIndex).Class Then Exit Function
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_LEFT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) - 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_LEFT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex) - 1) And (GetPlayerY(i) = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex) - 1) And (MapNpc(i).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_LEFT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If CheckMap(GetPlayerMap(MyIndex)).Left > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
        
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < MAX_MAPX Then
            ' Check for level block
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_LEVEL_BLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data1 >= Player(MyIndex).Level Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data2 <= Player(MyIndex).Level Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DirRight Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check for sex block
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SEX_BLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Sex Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DirRight Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
            
            ' Check to see if the map tile is blocked or not
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_BLOCKED Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_SIGN Then
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                        
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_CBLOCK Then
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data1 = Player(MyIndex).Class Then Exit Function
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data2 = Player(MyIndex).Class Then Exit Function
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Data3 = Player(MyIndex).Class Then Exit Function
                CanMove = False
                
                ' Set the new direction if they weren't facing that direction
                If d <> DIR_RIGHT Then
                    Call SendPlayerDir
                End If
                Exit Function
            End If
                                        
            ' Check to see if the key door is open or not
            If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_KEY Or CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).Type = TILE_TYPE_DOOR Then
                ' This actually checks if its open or not
                If TempTile(GetPlayerX(MyIndex) + 1, GetPlayerY(MyIndex)).DoorOpen = NO Then
                    CanMove = False
                    
                    ' Set the new direction if they weren't facing that direction
                    If d <> DIR_RIGHT Then
                        Call SendPlayerDir
                    End If
                    Exit Function
                End If
            End If
                        
            ' Check to see if a player is already on that tile
            For i = 1 To MAX_PLAYERS
                If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                    If (GetPlayerX(i) = GetPlayerX(MyIndex) + 1) And (GetPlayerY(i) = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_RIGHT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        
            ' Check to see if a npc is already on that tile
            For i = 1 To MAX_MAP_NPCS
                If MapNpc(i).Num > 0 Then
                    If (MapNpc(i).x = GetPlayerX(MyIndex) + 1) And (MapNpc(i).y = GetPlayerY(MyIndex)) Then
                        CanMove = False
                        
                        ' Set the new direction if they weren't facing that direction
                        If d <> DIR_RIGHT Then
                            Call SendPlayerDir
                        End If
                        Exit Function
                    End If
                End If
            Next i
        Else
            ' Check if they can warp to a new map
            If CheckMap(GetPlayerMap(MyIndex)).Right > 0 Then
                Call SendPlayerRequestNewMap
                GettingMap = True
            End If
            CanMove = False
            Exit Function
        End If
    End If
End Function

Sub CheckMovement()
    If GettingMap = False Then
        If IsTryingToMove Then
            If CanMove Then
                ' Check if player has the shift key down for running
                If ShiftDown Then
                    Player(MyIndex).Moving = MOVING_RUNNING
                Else
                    Player(MyIndex).Moving = MOVING_WALKING
                End If
            
                Select Case GetPlayerDir(MyIndex)
                    Case DIR_UP
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) - 1)
                
                    Case DIR_DOWN
                        Call SendPlayerMove
                        Player(MyIndex).YOffset = PIC_Y * -1
                        Call SetPlayerY(MyIndex, GetPlayerY(MyIndex) + 1)
                
                    Case DIR_LEFT
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) - 1)
                
                    Case DIR_RIGHT
                        Call SendPlayerMove
                        Player(MyIndex).XOffset = PIC_X * -1
                        Call SetPlayerX(MyIndex, GetPlayerX(MyIndex) + 1)
                End Select
            
                ' Gotta check :)
                If CheckMap(GetPlayerMap(MyIndex)).Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                    GettingMap = True
                End If
            End If
        End If
    End If
End Sub

Function FindPlayer(ByVal name As String) As Long
Dim i As Long

    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name thats to small
            If Len(GetPlayerName(i)) >= Len(Trim(name)) Then
                If UCase(Mid(GetPlayerName(i), 1, Len(Trim(name)))) = UCase(Trim(name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next i
    
    FindPlayer = 0
End Function

Public Sub EditorInit()
    SaveMap = CheckMap(GetPlayerMap(MyIndex))
    InEditor = True
    frmMapEditor.Show vbModeless, frmMirage
    'frmMirage.picMapEditor.Visible = True
    'With frmMapEditor.picBackSelect
        '.Width = 14 * PIC_X
        '.Height = 16384
        '.Picture = LoadPicture(App.Path + "\GFX\tiles.bmp")
    'End With
    frmMapEditor.picBackSelect.Picture = LoadPicture(App.Path + "\GFX\tiles.bmp")
    frmMapEditor.MouseSelected.Picture = LoadPicture(App.Path + "\GFX\tiles.bmp")
End Sub

Public Sub EditorMouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1, y1 As Long
Dim x2 As Long, y2 As Long, x3 As Long, y3 As Long

    If InEditor Then
        x1 = Int(x / PIC_X)
        y1 = Int(y / PIC_Y)
        If (Button = 1) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmMapEditor.shpSelected.Height <= PIC_Y And frmMapEditor.shpSelected.Width <= PIC_X Then
                If frmMapEditor.optLayers.Value = True Then
                    With CheckMap(GetPlayerMap(MyIndex)).Tile(x1, y1)
                        If frmMapEditor.optGround.Value = True Then .Ground = EditorTileY * 14 + EditorTileX
                            If frmMapEditor.optMask.Value = True Then .Mask = EditorTileY * 14 + EditorTileX
                            If frmMapEditor.optAnim.Value = True Then .Anim = EditorTileY * 14 + EditorTileX
                            If frmMapEditor.optMask2.Value = True Then .Mask2 = EditorTileY * 14 + EditorTileX
                            If frmMapEditor.optM2Anim.Value = True Then .M2Anim = EditorTileY * 14 + EditorTileX
                            If frmMapEditor.optFringe.Value = True Then .Fringe = EditorTileY * 14 + EditorTileX
                            If frmMapEditor.optFAnim.Value = True Then .FAnim = EditorTileY * 14 + EditorTileX
                            If frmMapEditor.optFringe2.Value = True Then .Fringe2 = EditorTileY * 14 + EditorTileX
                            If frmMapEditor.optF2Anim.Value = True Then .F2Anim = EditorTileY * 14 + EditorTileX
                    End With
                Else
                    With CheckMap(GetPlayerMap(MyIndex)).Tile(x1, y1)
                        If frmMapEditor.optBlocked.Value = True Then .Type = TILE_TYPE_BLOCKED
                        If frmMapEditor.optWarp.Value = True Then
                            .Type = TILE_TYPE_WARP
                            .Data1 = EditorWarpMap
                            .Data2 = EditorWarpX
                            .Data3 = EditorWarpY
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
    
                        If frmMapEditor.optHeal.Value = True Then
                            .Type = TILE_TYPE_HEAL
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
    
                        If frmMapEditor.optKill.Value = True Then
                            .Type = TILE_TYPE_KILL
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
    
                        If frmMapEditor.optItem.Value = True Then
                            .Type = TILE_TYPE_ITEM
                            .Data1 = ItemEditorNum
                            .Data2 = ItemEditorValue
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optNpcAvoid.Value = True Then
                            .Type = TILE_TYPE_NPCAVOID
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optKey.Value = True Then
                            .Type = TILE_TYPE_KEY
                            .Data1 = KeyEditorNum
                            .Data2 = KeyEditorTake
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optKeyOpen.Value = True Then
                            .Type = TILE_TYPE_KEYOPEN
                            .Data1 = KeyOpenEditorX
                            .Data2 = KeyOpenEditorY
                            .Data3 = 0
                            .String1 = KeyOpenEditorMsg
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optShop.Value = True Then
                            .Type = TILE_TYPE_SHOP
                            .Data1 = EditorShopNum
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optCBlock.Value = True Then
                            .Type = TILE_TYPE_CBLOCK
                            .Data1 = EditorItemNum1
                            .Data2 = EditorItemNum2
                            .Data3 = EditorItemNum3
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optArena.Value = True Then
                            .Type = TILE_TYPE_ARENA
                            .Data1 = Arena1
                            .Data2 = Arena2
                            .Data3 = Arena3
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optSound.Value = True Then
                            .Type = TILE_TYPE_SOUND
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = SoundFileName
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optSprite.Value = True Then
                            .Type = TILE_TYPE_SPRITE_CHANGE
                            .Data1 = SpritePic
                            .Data2 = SpriteItem
                            .Data3 = SpritePrice
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optSign.Value = True Then
                            .Type = TILE_TYPE_SIGN
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = SignLine1
                            .String2 = SignLine2
                            .String3 = SignLine3
                        End If
                        If frmMapEditor.optDoor.Value = True Then
                            .Type = TILE_TYPE_DOOR
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optNotice.Value = True Then
                            .Type = TILE_TYPE_NOTICE
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = NoticeTitle
                            .String2 = NoticeText
                            .String3 = NoticeSound
                        End If
                        If frmMapEditor.optChest.Value = True Then
                            .Type = TILE_TYPE_CHEST
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optClassChange.Value = True Then
                            .Type = TILE_TYPE_CLASS_CHANGE
                            .Data1 = ClassChange
                            .Data2 = ClassChangeReq
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optScripted.Value = True Then
                            .Type = TILE_TYPE_SCRIPTED
                            .Data1 = ScriptNum
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optSexBlock.Value = True Then
                            .Type = TILE_TYPE_SEX_BLOCK
                            .Data1 = SexBlockNum
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optLevelBlock.Value = True Then
                            .Type = TILE_TYPE_LEVEL_BLOCK
                            .Data1 = LevelBlockLow
                            .Data2 = LevelBlockHigh
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                        If frmMapEditor.optBank.Value = True Then
                            .Type = TILE_TYPE_BANK
                            .Data1 = 0
                            .Data2 = 0
                            .Data3 = 0
                            .String1 = ""
                            .String2 = ""
                            .String3 = ""
                        End If
                    End With
                End If
            Else
                For y2 = 0 To Int(frmMapEditor.shpSelected.Height / PIC_Y) - 1
                    For x2 = 0 To Int(frmMapEditor.shpSelected.Width / PIC_X) - 1
                        If x1 + x2 <= MAX_MAPX Then
                            If y1 + y2 <= MAX_MAPY Then
                                If frmMapEditor.optLayers.Value = True Then
                                    With CheckMap(GetPlayerMap(MyIndex)).Tile(x1 + x2, y1 + y2)
                                        If frmMapEditor.optGround.Value = True Then .Ground = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                        If frmMapEditor.optMask.Value = True Then .Mask = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                        If frmMapEditor.optAnim.Value = True Then .Anim = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                        If frmMapEditor.optMask2.Value = True Then .Mask2 = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                        If frmMapEditor.optM2Anim.Value = True Then .M2Anim = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                        If frmMapEditor.optFringe.Value = True Then .Fringe = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                        If frmMapEditor.optFAnim.Value = True Then .FAnim = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                        If frmMapEditor.optFringe2.Value = True Then .Fringe2 = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                        If frmMapEditor.optF2Anim.Value = True Then .F2Anim = (EditorTileY + y2) * 14 + (EditorTileX + x2)
                                    End With
                                End If
                            End If
                        End If
                    Next x2
                Next y2
            End If
            End If
        
        If (Button = 2) And (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
            If frmMapEditor.optLayers.Value = True Then
                With CheckMap(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    If frmMapEditor.optGround.Value = True Then .Ground = 0
                    If frmMapEditor.optMask.Value = True Then .Mask = 0
                    If frmMapEditor.optAnim.Value = True Then .Anim = 0
                    If frmMapEditor.optMask2.Value = True Then .Mask2 = 0
                    If frmMapEditor.optM2Anim.Value = True Then .M2Anim = 0
                    If frmMapEditor.optFringe.Value = True Then .Fringe = 0
                    If frmMapEditor.optFAnim.Value = True Then .FAnim = 0
                    If frmMapEditor.optFringe2.Value = True Then .Fringe2 = 0
                    If frmMapEditor.optF2Anim.Value = True Then .F2Anim = 0
                End With
            Else
                With CheckMap(GetPlayerMap(MyIndex)).Tile(x1, y1)
                    .Type = 0
                    .Data1 = 0
                    .Data2 = 0
                    .Data3 = 0
                    .String1 = ""
                    .String2 = ""
                    .String3 = ""
                End With
            End If
        End If
    End If
End Sub

Public Sub EditorChooseTile(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        EditorTileX = Int(x / PIC_X)
        EditorTileY = Int(y / PIC_Y)
    End If
    frmMapEditor.shpSelected.Top = Int(EditorTileY * PIC_Y)
    frmMapEditor.shpSelected.Left = Int(EditorTileX * PIC_Y)
    'Call BitBlt(frmMapEditor.picSelect.hDC, 0, 0, PIC_X, PIC_Y, frmMapEditor.picBackSelect.hDC, EditorTileX * PIC_X, EditorTileY * PIC_Y, SRCCOPY)
End Sub

Public Sub EditorTileScroll()
    frmMapEditor.picBackSelect.Top = (frmMapEditor.scrlPicture.Value * PIC_Y) * -1
End Sub

Public Sub EditorSend()
    Call SendMap
    Call EditorCancel
End Sub

Public Sub EditorCancel()
    InEditor = False
    frmMapEditor.Visible = False
    'frmmirage.show
    'frmMirage.picMapEditor.Visible = False
End Sub

Public Sub EditorClearLayer()
Dim YesNo As Long, x As Long, y As Long

    ' Ground layer
    If frmMapEditor.optGround.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the ground layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Ground = 0
                Next x
            Next y
        End If
    End If

    ' Mask layer
    If frmMapEditor.optMask.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Mask = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask Animation layer
    If frmMapEditor.optAnim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Anim = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask 2 layer
    If frmMapEditor.optMask2.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Mask2 = 0
                Next x
            Next y
        End If
    End If
    
    ' Mask 2 Animation layer
    If frmMapEditor.optM2Anim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the mask 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).M2Anim = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe layer
    If frmMapEditor.optFringe.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Fringe = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe Animation layer
    If frmMapEditor.optFAnim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).FAnim = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe 2 layer
    If frmMapEditor.optFringe2.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Fringe2 = 0
                Next x
            Next y
        End If
    End If
    
    ' Fringe 2 Animation layer
    If frmMapEditor.optF2Anim.Value = True Then
    YesNo = MsgBox("Are you sure you wish to clear the fringe 2 animation layer?", vbYesNo, GAME_NAME)
        If YesNo = vbYes Then
            For y = 0 To MAX_MAPY
                For x = 0 To MAX_MAPX
                    CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).F2Anim = 0
                Next x
            Next y
        End If
    End If
End Sub

Public Sub EditorClearAttribs()
Dim YesNo As Long, x As Long, y As Long

    YesNo = MsgBox("Are you sure you wish to clear the attributes on this map?", vbYesNo, GAME_NAME)
    
    If YesNo = vbYes Then
        For y = 0 To MAX_MAPY
            For x = 0 To MAX_MAPX
                CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Type = 0
            Next x
        Next y
    End If
End Sub

Public Sub EmoticonEditorInit()
    frmEmoticonEditor.scrlEmoticon.Value = Emoticons(EditorIndex - 1).Pic
    frmEmoticonEditor.txtCommand.Text = Trim(Emoticons(EditorIndex - 1).Command)
    frmEmoticonEditor.picEmoticons.Picture = LoadPicture(App.Path & "\GFX\emoticons.bmp")
    frmEmoticonEditor.Show vbModal
End Sub

Public Sub EmoticonEditorOk()
    Emoticons(EditorIndex - 1).Pic = frmEmoticonEditor.scrlEmoticon.Value
    If frmEmoticonEditor.txtCommand.Text <> "/" Then
        Emoticons(EditorIndex - 1).Command = frmEmoticonEditor.txtCommand.Text
    Else
        Emoticons(EditorIndex - 1).Command = ""
    End If
    
    Call SendSaveEmoticon(EditorIndex - 1)
    Call EmoticonEditorCancel
End Sub

Public Sub EmoticonEditorCancel()
    InEmoticonEditor = False
    Unload frmEmoticonEditor
End Sub

Public Sub ItemEditorInit()

    EditorItemY = Int(Item(EditorIndex).Pic / 6)
    EditorItemX = (Item(EditorIndex).Pic - Int(Item(EditorIndex).Pic / 6) * 6)
    
    frmItemEditor.scrlClassReq.Max = Max_Classes

    frmItemEditor.picItems.Picture = LoadPicture(App.Path & "\GFX\items.bmp")
    
    frmItemEditor.txtName.Text = Trim(Item(EditorIndex).name)
    frmItemEditor.txtDesc.Text = Trim(Item(EditorIndex).desc)
    frmItemEditor.cmbType.ListIndex = Item(EditorIndex).Type
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Or (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_BOOTS) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_GLOVES) Then
        frmItemEditor.fraEquipment.Visible = True
        frmItemEditor.scrlDurability.Value = Item(EditorIndex).Data1
        frmItemEditor.scrlStrength.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlStrReq.Value = Item(EditorIndex).StrReq
        frmItemEditor.scrlDefReq.Value = Item(EditorIndex).DefReq
        frmItemEditor.scrlSpeedReq.Value = Item(EditorIndex).SpeedReq
        frmItemEditor.scrlClassReq.Value = Item(EditorIndex).ClassReq
        frmItemEditor.scrlAccessReq.Value = Item(EditorIndex).AccessReq
        frmItemEditor.scrlAddHP.Value = Item(EditorIndex).AddHP
        frmItemEditor.scrlAddMP.Value = Item(EditorIndex).AddMP
        frmItemEditor.scrlAddSP.Value = Item(EditorIndex).AddSP
        frmItemEditor.scrlAddStr.Value = Item(EditorIndex).AddStr
        frmItemEditor.scrlAddDef.Value = Item(EditorIndex).AddDef
        frmItemEditor.scrlAddVit.Value = Item(EditorIndex).Data3
        frmItemEditor.scrlAddMagi.Value = Item(EditorIndex).AddMagi
        frmItemEditor.scrlAddSpeed.Value = Item(EditorIndex).AddSpeed
        frmItemEditor.scrlAddEXP.Value = Item(EditorIndex).AddEXP
    Else
        frmItemEditor.fraEquipment.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        frmItemEditor.fraVitals.Visible = True
        frmItemEditor.scrlVitalMod.Value = Item(EditorIndex).Data1
    Else
        frmItemEditor.fraVitals.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        frmItemEditor.fraSpell.Visible = True
        frmItemEditor.scrlSpell.Value = Item(EditorIndex).Data1
        
    Else
        frmItemEditor.fraSpell.Visible = False
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCROLL) Or (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_ORB) Then
        frmItemEditor.fraWarp.Visible = True
        frmItemEditor.txtMap.Text = Item(EditorIndex).Data1
        frmItemEditor.scrlX.Value = Item(EditorIndex).Data2
        frmItemEditor.scrlY.Value = Item(EditorIndex).Data3
    Else
        frmItemEditor.fraWarp.Visible = False
    End If
    
    frmItemEditor.Show vbModal
End Sub

Public Sub ItemEditorOk()
    Item(EditorIndex).name = frmItemEditor.txtName.Text
    Item(EditorIndex).desc = frmItemEditor.txtDesc.Text
    Item(EditorIndex).Pic = EditorItemY * 6 + EditorItemX
    Item(EditorIndex).Type = frmItemEditor.cmbType.ListIndex

    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_WEAPON) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_SHIELD) Or (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_BOOTS) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_GLOVES) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlDurability.Value
        Item(EditorIndex).Data2 = frmItemEditor.scrlStrength.Value
        Item(EditorIndex).Data3 = frmItemEditor.scrlAddVit.Value
        Item(EditorIndex).StrReq = frmItemEditor.scrlStrReq.Value
        Item(EditorIndex).DefReq = frmItemEditor.scrlDefReq.Value
        Item(EditorIndex).SpeedReq = frmItemEditor.scrlSpeedReq.Value
        
        Item(EditorIndex).ClassReq = frmItemEditor.scrlClassReq.Value
        Item(EditorIndex).AccessReq = frmItemEditor.scrlAccessReq.Value
        
        Item(EditorIndex).AddHP = frmItemEditor.scrlAddHP.Value
        Item(EditorIndex).AddMP = frmItemEditor.scrlAddMP.Value
        Item(EditorIndex).AddSP = frmItemEditor.scrlAddSP.Value
        Item(EditorIndex).AddStr = frmItemEditor.scrlAddStr.Value
        Item(EditorIndex).AddDef = frmItemEditor.scrlAddDef.Value
        Item(EditorIndex).AddMagi = frmItemEditor.scrlAddMagi.Value
        Item(EditorIndex).AddSpeed = frmItemEditor.scrlAddSpeed.Value
        Item(EditorIndex).AddEXP = frmItemEditor.scrlAddEXP.Value
    End If
    
    If (frmItemEditor.cmbType.ListIndex >= ITEM_TYPE_POTIONADDHP) And (frmItemEditor.cmbType.ListIndex <= ITEM_TYPE_POTIONSUBSP) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlVitalMod.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SPELL) Then
        Item(EditorIndex).Data1 = frmItemEditor.scrlSpell.Value
        Item(EditorIndex).Data2 = 0
        Item(EditorIndex).Data3 = 0
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
    End If
    
    If (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_SCROLL) Or (frmItemEditor.cmbType.ListIndex = ITEM_TYPE_ORB) Then
        If IsNumeric(frmItemEditor.txtMap.Text) = False Then
            MsgBox "Please input numbers only!"
            Exit Sub
        End If
        Item(EditorIndex).Data1 = Val(frmItemEditor.txtMap.Text)
        Item(EditorIndex).Data2 = frmItemEditor.scrlX.Value
        Item(EditorIndex).Data3 = frmItemEditor.scrlY.Value
        Item(EditorIndex).StrReq = 0
        Item(EditorIndex).DefReq = 0
        Item(EditorIndex).SpeedReq = 0
        Item(EditorIndex).ClassReq = -1
        Item(EditorIndex).AccessReq = 0
        
        Item(EditorIndex).AddHP = 0
        Item(EditorIndex).AddMP = 0
        Item(EditorIndex).AddSP = 0
        Item(EditorIndex).AddStr = 0
        Item(EditorIndex).AddDef = 0
        Item(EditorIndex).AddMagi = 0
        Item(EditorIndex).AddSpeed = 0
        Item(EditorIndex).AddEXP = 0
    End If
    
    Call SendSaveItem(EditorIndex)
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub ItemEditorCancel()
    InItemsEditor = False
    Unload frmItemEditor
End Sub

Public Sub NpcEditorInit()
    
    frmNpcEditor.picSprites.Picture = LoadPicture(App.Path & "\GFX\sprites.bmp")
    
    frmNpcEditor.txtName.Text = Trim(Npc(EditorIndex).name)
    frmNpcEditor.txtAttackSay.Text = Trim(Npc(EditorIndex).AttackSay)
    frmNpcEditor.scrlSprite.Value = Npc(EditorIndex).Sprite
    frmNpcEditor.txtSpawnSecs.Text = STR(Npc(EditorIndex).SpawnSecs)
    frmNpcEditor.cmbBehavior.ListIndex = Npc(EditorIndex).Behavior
    frmNpcEditor.scrlRange.Value = Npc(EditorIndex).Range
    frmNpcEditor.scrlSTR.Value = Npc(EditorIndex).STR
    If Npc(EditorIndex).DEF > frmNpcEditor.scrlDEF.Max Then Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Max
    frmNpcEditor.scrlDEF.Value = Npc(EditorIndex).DEF
    If Npc(EditorIndex).speed > frmNpcEditor.scrlSPEED.Max Then Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.Max
    frmNpcEditor.scrlSPEED.Value = Npc(EditorIndex).speed
    frmNpcEditor.scrlMAGI.Value = Npc(EditorIndex).MAGI
    frmNpcEditor.BigNpc.Value = Npc(EditorIndex).Big
    frmNpcEditor.StartHP.Value = Npc(EditorIndex).MaxHp
    frmNpcEditor.ExpGive.Value = Npc(EditorIndex).EXP
    frmNpcEditor.txtChance.Text = STR(Npc(EditorIndex).ItemNPC(1).Chance)
    frmNpcEditor.scrlNum.Value = Npc(EditorIndex).ItemNPC(1).ItemNum
    frmNpcEditor.scrlValue.Value = Npc(EditorIndex).ItemNPC(1).ItemValue
    
    frmNpcEditor.Show vbModal
End Sub

Public Sub NpcEditorOk()
    Npc(EditorIndex).name = frmNpcEditor.txtName.Text
    Npc(EditorIndex).AttackSay = frmNpcEditor.txtAttackSay.Text
    Npc(EditorIndex).Sprite = frmNpcEditor.scrlSprite.Value
    Npc(EditorIndex).SpawnSecs = Val(frmNpcEditor.txtSpawnSecs.Text)
    Npc(EditorIndex).Behavior = frmNpcEditor.cmbBehavior.ListIndex
    Npc(EditorIndex).Range = frmNpcEditor.scrlRange.Value
    Npc(EditorIndex).STR = frmNpcEditor.scrlSTR.Value
    Npc(EditorIndex).DEF = frmNpcEditor.scrlDEF.Value
    Npc(EditorIndex).speed = frmNpcEditor.scrlSPEED.Value
    Npc(EditorIndex).MAGI = frmNpcEditor.scrlMAGI.Value
    Npc(EditorIndex).Big = frmNpcEditor.BigNpc.Value
    Npc(EditorIndex).MaxHp = frmNpcEditor.StartHP.Value
    Npc(EditorIndex).EXP = frmNpcEditor.ExpGive.Value
    
    Call SendSaveNpc(EditorIndex)
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorCancel()
    InNpcEditor = False
    Unload frmNpcEditor
End Sub

Public Sub NpcEditorBltSprite()
On Error Resume Next
    If frmNpcEditor.BigNpc.Value = Checked Then
        frmNpcEditor.picSprites.Left = -(3 * 64)
        frmNpcEditor.picSprites.Top = -(frmNpcEditor.scrlSprite.Value * 64)
    Else
        frmNpcEditor.picSprites.Left = -(3 * 32)
        frmNpcEditor.picSprites.Top = -(frmNpcEditor.scrlSprite.Value * PIC_Y)
    End If
End Sub

Public Sub ShopEditorInit()
Dim i As Long

    frmShopEditor.txtName.Text = Trim(Shop(EditorIndex).name)
    frmShopEditor.txtJoinSay.Text = Trim(Shop(EditorIndex).JoinSay)
    frmShopEditor.txtLeaveSay.Text = Trim(Shop(EditorIndex).LeaveSay)
    frmShopEditor.chkFixesItems.Value = Shop(EditorIndex).FixesItems
    
    frmShopEditor.cmbItemGive.Clear
    frmShopEditor.cmbItemGive.AddItem "None"
    frmShopEditor.cmbItemGet.Clear
    frmShopEditor.cmbItemGet.AddItem "None"
    For i = 1 To MAX_ITEMS
        frmShopEditor.cmbItemGive.AddItem i & ": " & Trim(Item(i).name)
        frmShopEditor.cmbItemGet.AddItem i & ": " & Trim(Item(i).name)
    Next i
    frmShopEditor.cmbItemGive.ListIndex = 0
    frmShopEditor.cmbItemGet.ListIndex = 0
    
    Call UpdateShopTrade
    
    frmShopEditor.Show vbModal
End Sub

Public Sub UpdateShopTrade()
Dim i As Long, GetItem As Long, GetValue As Long, GiveItem As Long, GiveValue As Long
    
    frmShopEditor.lstTradeItem.Clear
    For i = 1 To MAX_TRADES
        GetItem = Shop(EditorIndex).TradeItem(i).GetItem
        GetValue = Shop(EditorIndex).TradeItem(i).GetValue
        GiveItem = Shop(EditorIndex).TradeItem(i).GiveItem
        GiveValue = Shop(EditorIndex).TradeItem(i).GiveValue
        
        If GetItem > 0 And GiveItem > 0 Then
            frmShopEditor.lstTradeItem.AddItem i & ": " & GiveValue & " " & Trim(Item(GiveItem).name) & " for " & GetValue & " " & Trim(Item(GetItem).name)
        Else
            frmShopEditor.lstTradeItem.AddItem "Empty Trade Slot"
        End If
    Next i
    frmShopEditor.lstTradeItem.ListIndex = 0
End Sub

Public Sub ShopEditorOk()
    Shop(EditorIndex).name = frmShopEditor.txtName.Text
    Shop(EditorIndex).JoinSay = frmShopEditor.txtJoinSay.Text
    Shop(EditorIndex).LeaveSay = frmShopEditor.txtLeaveSay.Text
    Shop(EditorIndex).FixesItems = frmShopEditor.chkFixesItems.Value
    
    Call SendSaveShop(EditorIndex)
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub ShopEditorCancel()
    InShopEditor = False
    Unload frmShopEditor
End Sub

Public Sub SpellEditorInit()
Dim i As Long

    frmSpellEditor.cmbClassReq.AddItem "All Classes"
    For i = 0 To Max_Classes
        frmSpellEditor.cmbClassReq.AddItem Trim(Class(i).name)
    Next i
    
    frmSpellEditor.txtName.Text = Trim(Spell(EditorIndex).name)
    frmSpellEditor.cmbClassReq.ListIndex = Spell(EditorIndex).ClassReq
    frmSpellEditor.scrlLevelReq.Value = Spell(EditorIndex).LevelReq
        
    frmSpellEditor.cmbType.ListIndex = Spell(EditorIndex).Type
    
    If Spell(EditorIndex).Data1 < frmSpellEditor.scrlVitalMod.Min Then Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.Min
    frmSpellEditor.scrlVitalMod.Value = Spell(EditorIndex).Data1
    
    frmSpellEditor.scrlCost.Value = Spell(EditorIndex).MPCost
    frmSpellEditor.scrlSound.Value = Spell(EditorIndex).Sound
        
    frmSpellEditor.Show vbModal
End Sub

Public Sub SpellEditorOk()
    Spell(EditorIndex).name = frmSpellEditor.txtName.Text
    Spell(EditorIndex).ClassReq = frmSpellEditor.cmbClassReq.ListIndex
    Spell(EditorIndex).LevelReq = frmSpellEditor.scrlLevelReq.Value
    Spell(EditorIndex).Type = frmSpellEditor.cmbType.ListIndex
    Spell(EditorIndex).Data1 = frmSpellEditor.scrlVitalMod.Value
    Spell(EditorIndex).Data3 = 0
    Spell(EditorIndex).MPCost = frmSpellEditor.scrlCost.Value
    Spell(EditorIndex).Sound = frmSpellEditor.scrlSound.Value
    
    Call SendSaveSpell(EditorIndex)
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub SpellEditorCancel()
    InSpellEditor = False
    Unload frmSpellEditor
End Sub

Public Sub UpdateTradeInventory()
Dim i As Long

    frmPlayerTrade.PlayerInv1.Clear
    
For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
            Else
                frmPlayerTrade.PlayerInv1.AddItem i & ": " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name)
            End If
        End If
    Else
        frmPlayerTrade.PlayerInv1.AddItem "<Nothing>"
    End If
Next i
    
    frmPlayerTrade.PlayerInv1.ListIndex = 0
End Sub

Sub PlayerSearch(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim x1 As Long, y1 As Long

    x1 = Int(x / PIC_X)
    y1 = Int(y / PIC_Y)
    
    If (x1 >= 0) And (x1 <= MAX_MAPX) And (y1 >= 0) And (y1 <= MAX_MAPY) Then
        Call SendData("search" & SEP_CHAR & x1 & SEP_CHAR & y1 & SEP_CHAR & END_CHAR)
    End If
    MouseDownX = x1
    MouseDownY = y1
End Sub
Sub BltTile2(ByVal x As Long, ByVal y As Long, ByVal Tile As Long)
    rec.Top = Int(Tile / 14) * PIC_Y
    rec.Bottom = rec.Top + PIC_Y
    rec.Left = (Tile - Int(Tile / 14) * 14) * PIC_X
    rec.Right = rec.Left + PIC_X
    Call DD_BackBuffer.BltFast((PIC_X * (MAX_MAPX + 1)) + (x - (NewPlayerX * PIC_X) + sx - NewXOffset), (PIC_Y * (MAX_MAPY + 1)) + (y - (NewPlayerY * PIC_Y) + sx - NewYOffset), DD_TileSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
End Sub

Sub BltPlayerText(ByVal Index As Long)
Dim TextX As Long
Dim TextY As Long
Dim intLoop As Integer
Dim intLoop2 As Integer

Dim bytLineCount As Byte
Dim bytLineLength As Byte
Dim strLine(0 To MAX_LINES - 1) As String
Dim strWords() As String
    strWords() = Split(Bubble(Index).Text, " ")
    
    If Len(Bubble(Index).Text) < MAX_LINE_LENGTH Then
        DISPLAY_BUBBLE_WIDTH = 2 + ((Len(Bubble(Index).Text) * 9) \ PIC_X)
        
        If DISPLAY_BUBBLE_WIDTH > MAX_BUBBLE_WIDTH Then
            DISPLAY_BUBBLE_WIDTH = MAX_BUBBLE_WIDTH
        End If
    Else
        DISPLAY_BUBBLE_WIDTH = 6
    End If
    
    TextX = GetPlayerX(Index) * PIC_X + Player(Index).XOffset + Int(PIC_X) - ((DISPLAY_BUBBLE_WIDTH * 32) / 2) - 6
    TextY = GetPlayerY(Index) * PIC_Y + Player(Index).YOffset - Int(PIC_Y) + 85
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    ' Draw the fancy box with tiles.
    Call BltTile2(TextX - 10, TextY - 10, 2)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY - 10, 3)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY - 10, 17)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    ' Loop through all the words.
    For intLoop = 0 To UBound(strWords)
        ' Increment the line length.
        bytLineLength = bytLineLength + Len(strWords(intLoop)) + 1
            
        ' If we have room on the current line.
        If bytLineLength < MAX_LINE_LENGTH Then
            ' Add the text to the current line.
            strLine(bytLineCount) = strLine(bytLineCount) & strWords(intLoop) & " "
        Else
            bytLineCount = bytLineCount + 1
            
            If bytLineCount = MAX_LINES Then
                bytLineCount = bytLineCount - 1
                Exit For
            End If
            
            strLine(bytLineCount) = Trim(strWords(intLoop)) & " "
            bytLineLength = 0
        End If
    Next intLoop
    
    Call DD_BackBuffer.ReleaseDC(TexthDC)
    
    If bytLineCount > 0 Then
        For intLoop = 6 To (bytLineCount - 2) * PIC_Y + 6
            Call BltTile2(TextX - 10, TextY - 10 + intLoop, 19)
            Call BltTile2(TextX - 10 + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X, TextY - 10 + intLoop, 18)
            
            For intLoop2 = 1 To DISPLAY_BUBBLE_WIDTH - 2
                Call BltTile2(TextX - 10 + (intLoop2 * PIC_X), TextY + intLoop - 10, 6)
            Next intLoop2
        Next intLoop
    End If

    Call BltTile2(TextX - 10, TextY + (bytLineCount * 16) - 4, 4)
    Call BltTile2(TextX + (DISPLAY_BUBBLE_WIDTH * PIC_X) - PIC_X - 10, TextY + (bytLineCount * 16) - 4, 5)
    
    For intLoop = 1 To (DISPLAY_BUBBLE_WIDTH - 2)
        Call BltTile2(TextX - 10 + (intLoop * PIC_X), TextY + (bytLineCount * 16) - 4, 16)
    Next intLoop
    
    TexthDC = DD_BackBuffer.GetDC
    
    For intLoop = 0 To (MAX_LINES - 1)
        If strLine(intLoop) <> "" Then
            Call DrawText(TexthDC, TextX - (NewPlayerX * PIC_X) + sx - NewXOffset + (((DISPLAY_BUBBLE_WIDTH) * PIC_X) / 2) - ((Len(strLine(intLoop)) * GameFontSize) \ 2) - 7, TextY - (NewPlayerY * PIC_Y) + sx - NewYOffset, strLine(intLoop), QBColor(DarkGrey))
            TextY = TextY + 16
        End If
    Next intLoop
End Sub
Sub BltPlayerBar()
Dim x As Long, y As Long

If Player(MyIndex).HP <= 0 Then Exit Sub

    x = (PIC_X * (MAX_MAPX + 1)) + ((GetPlayerX(MyIndex) * PIC_X + sx + Player(MyIndex).XOffset) - (NewPlayerX * PIC_X) - NewXOffset)
    y = (PIC_Y * (MAX_MAPY + 1)) + ((GetPlayerY(MyIndex) * PIC_Y + sx + Player(MyIndex).YOffset) - (NewPlayerY * PIC_Y) - NewYOffset)

    'draws the back bars
    Call DD_BackBuffer.SetFillColor(RGB(255, 0, 0))
    Call DD_BackBuffer.DrawBox(x, y + 32, x + 32, y + 36)
    
    'draws HP
    Call DD_BackBuffer.SetFillColor(RGB(0, 255, 0))
    Call DD_BackBuffer.DrawBox(x, y + 32, x + ((Player(MyIndex).HP / 100) / (Player(MyIndex).MaxHp / 100) * 32), y + 36)
End Sub

Public Sub UpdateVisInv()
Dim Index As Long
Dim d As Long
Dim c As Long

c = 0
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) <> Index Then frmMirage.ShieldImage.Picture = LoadPicture()
        If GetPlayerWeaponSlot(MyIndex) <> Index Then frmMirage.WeaponImage.Picture = LoadPicture()
        If GetPlayerHelmetSlot(MyIndex) <> Index Then frmMirage.HelmetImage.Picture = LoadPicture()
        If GetPlayerArmorSlot(MyIndex) <> Index Then frmMirage.ArmorImage.Picture = LoadPicture()
        If GetPlayerBootsSlot(MyIndex) <> Index Then frmMirage.BootsImage.Picture = LoadPicture()
        If GetPlayerGlovesSlot(MyIndex) <> Index Then frmMirage.GlovesImage.Picture = LoadPicture()
        If GetPlayerRingSlot(MyIndex) <> Index Then frmMirage.RingImage.Picture = LoadPicture()
        If GetPlayerAmuletSlot(MyIndex) <> Index Then frmMirage.AmuletImage.Picture = LoadPicture()
    
    Next Index
    
    For Index = 1 To MAX_INV
        If GetPlayerShieldSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ShieldImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerWeaponSlot(MyIndex) = Index Then Call BitBlt(frmMirage.WeaponImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerHelmetSlot(MyIndex) = Index Then Call BitBlt(frmMirage.HelmetImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerArmorSlot(MyIndex) = Index Then Call BitBlt(frmMirage.ArmorImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerBootsSlot(MyIndex) = Index Then Call BitBlt(frmMirage.BootsImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerGlovesSlot(MyIndex) = Index Then Call BitBlt(frmMirage.GlovesImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerRingSlot(MyIndex) = Index Then Call BitBlt(frmMirage.RingImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        If GetPlayerAmuletSlot(MyIndex) = Index Then Call BitBlt(frmMirage.AmuletImage.hDC, 0, 0, PIC_X, PIC_Y, frmMirage.picItems.hDC, (Item(GetPlayerInvItemNum(MyIndex, Index)).Pic - Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * 6) * PIC_X, Int(Item(GetPlayerInvItemNum(MyIndex, Index)).Pic / 6) * PIC_Y, SRCCOPY)
        
    Next Index
        
    frmMirage.EquipS(0).Visible = False
    frmMirage.EquipS(1).Visible = False
    frmMirage.EquipS(2).Visible = False
    frmMirage.EquipS(3).Visible = False
    frmMirage.EquipS(4).Visible = False
    frmMirage.EquipS(5).Visible = False
    frmMirage.EquipS(6).Visible = False
    frmMirage.EquipS(7).Visible = False

    For d = 0 To MAX_INV - 1
        If Player(MyIndex).Inv(d + 1).Num > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type <> ITEM_TYPE_CURRENCY Then
                If GetPlayerWeaponSlot(MyIndex) = d + 1 Then
                    frmMirage.EquipS(0).Visible = True
                    frmMirage.EquipS(0).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(0).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerArmorSlot(MyIndex) = d + 1 Then
                    frmMirage.EquipS(1).Visible = True
                    frmMirage.EquipS(1).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(1).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerHelmetSlot(MyIndex) = d + 1 Then
                    frmMirage.EquipS(2).Visible = True
                    frmMirage.EquipS(2).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(2).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerShieldSlot(MyIndex) = d + 1 Then
                    frmMirage.EquipS(3).Visible = True
                    frmMirage.EquipS(3).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(3).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerBootsSlot(MyIndex) = d + 1 Then
                    frmMirage.EquipS(4).Visible = True
                    frmMirage.EquipS(4).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(4).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerGlovesSlot(MyIndex) = d + 1 Then
                    frmMirage.EquipS(5).Visible = True
                    frmMirage.EquipS(5).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(5).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerRingSlot(MyIndex) = d + 1 Then
                    frmMirage.EquipS(6).Visible = True
                    frmMirage.EquipS(6).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(6).Left = frmMirage.picInv(d).Left - 2
                ElseIf GetPlayerAmuletSlot(MyIndex) = d + 1 Then
                    frmMirage.EquipS(7).Visible = True
                    frmMirage.EquipS(7).Top = frmMirage.picInv(d).Top - 2
                    frmMirage.EquipS(7).Left = frmMirage.picInv(d).Left - 2
                End If
            End If
            If Item(GetPlayerInvItemNum(MyIndex, d + 1)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemNum(MyIndex, d + 1) > 0 Then
                    If GetPlayerInvItemNum(MyIndex, d + 1) = 1 Then
                        frmMirage.lblGold.Caption = GetPlayerInvItemValue(MyIndex, d + 1)
                    End If
                Else
                    frmMirage.lblGold.Caption = 0
                End If
            End If
        End If
    Next d
End Sub

Sub BltSpriteChange(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    
    If CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = x * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        rec.Top = CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * PIC_Y + 16
        rec.Bottom = rec.Top + PIC_Y - 16
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx)
        y2 = (PIC_Y * (MAX_MAPY + 1)) + (y * PIC_Y + sx)
                                           
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltSpriteChange2(ByVal x As Long, ByVal y As Long)
Dim x2 As Long, y2 As Long
    
    If CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Type = TILE_TYPE_SPRITE_CHANGE Then

        ' Only used if ever want to switch to blt rather then bltfast
        With rec_pos
            .Top = y * PIC_Y
            .Bottom = .Top + PIC_Y
            .Left = x * PIC_X
            .Right = .Left + PIC_X
        End With
                                        
        rec.Top = CheckMap(GetPlayerMap(MyIndex)).Tile(x, y).Data1 * PIC_Y
        rec.Bottom = rec.Top + PIC_Y - 16
        rec.Left = 96
        rec.Right = rec.Left + PIC_X
        
        x2 = (PIC_X * (MAX_MAPX + 1)) + (x * PIC_X + sx)
        y2 = (PIC_Y * (MAX_MAPY + 1)) + (y * PIC_Y + (sx / 2)) '- 16
               
        If y2 < 0 Then
            Exit Sub
        End If
                            
        'Call DD_BackBuffer.Blt(rec_pos, DD_SpriteSurf, rec, DDBLT_WAIT Or DDBLT_KEYSRC)
        Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_SpriteSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
    End If
End Sub

Sub BltEmoticons(ByVal Index As Long)
Dim x2 As Long, y2 As Long
   
    If Player(Index).Emoticon < 0 Then Exit Sub
    
    If Player(Index).EmoticonT + 5000 > GetTickCount Then
        rec.Top = Player(Index).Emoticon * PIC_Y
        rec.Bottom = rec.Top + PIC_Y
        rec.Left = 0
        rec.Right = PIC_X
        
        If Index = MyIndex Then
            x2 = (PIC_X * (MAX_MAPX + 1)) + (NewX + sx + 16)
            y2 = (PIC_Y * (MAX_MAPY + 1)) + (NewY + sx - 32)
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2, y2, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        Else
            x2 = (PIC_X * (MAX_MAPX + 1)) + (GetPlayerX(Index) * PIC_X + sx + Player(Index).XOffset + 16)
            y2 = (PIC_Y * (MAX_MAPY + 1)) + (GetPlayerY(Index) * PIC_Y + sx + Player(Index).YOffset - 32)
            
            If y2 < 0 Then
                Exit Sub
            End If
            
            Call DD_BackBuffer.BltFast(x2 - (NewPlayerX * PIC_X) - NewXOffset, y2 - (NewPlayerY * PIC_Y) - NewYOffset, DD_EmoticonSurf, rec, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY)
        End If
    End If
End Sub

Sub CheckInput3()
Dim x As Long, y As Long, b As Long
Dim CurrJoy As JOYINFO
joyGetPos 0, CurrJoy
x = CurrJoy.wXpos
y = CurrJoy.wYpos
b = CurrJoy.wButtons

    If GettingMap = False Then
        If JUpC = 1 Then
            If x = JUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            Else
                DirUp = False
            End If
        ElseIf JUpC = 2 Then
            If y = JUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            Else
                DirUp = False
            End If
        Else
            If b = JUp Then
                DirUp = True
                DirDown = False
                DirLeft = False
                DirRight = False
            Else
                DirUp = False
            End If
        End If
        
        If JDownC = 1 Then
            If x = JDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            Else
                DirDown = False
            End If
        ElseIf JDownC = 2 Then
            If y = JDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            Else
                DirDown = False
            End If
        Else
            If b = JDown Then
                DirUp = False
                DirDown = True
                DirLeft = False
                DirRight = False
            Else
                DirDown = False
            End If
        End If
        
        If JLeftC = 1 Then
            If x = JLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            Else
                DirLeft = False
            End If
        ElseIf JLeftC = 2 Then
            If y = JLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            Else
                DirLeft = False
            End If
        Else
            If b = JLeft Then
                DirUp = False
                DirDown = False
                DirLeft = True
                DirRight = False
            Else
                DirLeft = False
            End If
        End If
        
        If JRightC = 1 Then
            If x = JRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            Else
                DirRight = False
            End If
        ElseIf JRightC = 2 Then
            If y = JRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            Else
                DirRight = False
            End If
        Else
            If b = JRight Then
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = True
            Else
                DirRight = False
            End If
        End If

        If CurrJoy.wButtons = (JAttack + JRun + JEnter) Then
            ShiftDown = True
            Call CheckMapGetItem
            ControlDown = True
        Else
            If CurrJoy.wButtons = (JRun + JEnter) Then
                ShiftDown = True
                Call CheckMapGetItem
            Else
                If CurrJoy.wButtons = (JAttack + JRun) Then
                    ControlDown = True
                    ShiftDown = True
                Else
                    If CurrJoy.wButtons = (JAttack + JEnter) Then
                        Call CheckMapGetItem
                        ControlDown = True
                    Else
                        If CurrJoy.wButtons = JAttack Then
                            ControlDown = True
                        Else
                            ControlDown = False
                        End If
                        
                        If CurrJoy.wButtons = JRun Then
                            ShiftDown = True
                        Else
                            ShiftDown = False
                        End If
                        If CurrJoy.wButtons = JEnter Then
                            Call CheckMapGetItem
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Sub CheckInput2()
Dim x As Long, y As Long, b As Long
Dim CurrJoy As JOYINFO
joyGetPos 0, CurrJoy
x = CurrJoy.wXpos
y = CurrJoy.wYpos
b = CurrJoy.wButtons

    If GettingMap = False Then
        If y = JUp Then
            DirUp = True
            DirDown = False
            DirLeft = False
            DirRight = False
        Else
            'DirUp = False
        End If
        If y = JDown Then
            DirUp = False
            DirDown = True
            DirLeft = False
            DirRight = False
        Else
            'DirDown = False
        End If
        If x = JLeft Then
            DirUp = False
            DirDown = False
            DirLeft = True
            DirRight = False
        Else
            'DirLeft = False
        End If
        If x = JRight Then
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = True
        Else
            'DirRight = False
        End If
        If CurrJoy.wButtons = (JAttack + JRun + JEnter) Then
            ShiftDown = True
            Call CheckMapGetItem
            ControlDown = True
        Else
            If CurrJoy.wButtons = (JRun + JEnter) Then
                ShiftDown = True
                Call CheckMapGetItem
            Else
                If CurrJoy.wButtons = (JAttack + JRun) Then
                    ControlDown = True
                    ShiftDown = True
                Else
                    If CurrJoy.wButtons = (JAttack + JEnter) Then
                        Call CheckMapGetItem
                        ControlDown = True
                    Else
                        If CurrJoy.wButtons = JAttack Then
                            ControlDown = True
                        Else
                            'ControlDown = False
                        End If
                        
                        If CurrJoy.wButtons = JRun Then
                            ShiftDown = True
                        Else
                            'ShiftDown = False
                        End If
                        If CurrJoy.wButtons = JEnter Then
                            Call CheckMapGetItem
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Sub UpdateBank()
Dim i As Long

frmBank.lstInventory.Clear
frmBank.lstBank.Clear

For i = 1 To MAX_INV
    If GetPlayerInvItemNum(MyIndex, i) > 0 Then
        If Item(GetPlayerInvItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (" & GetPlayerInvItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name) & " (worn)"
            Else
                frmBank.lstInventory.AddItem i & "> " & Trim(Item(GetPlayerInvItemNum(MyIndex, i)).name)
            End If
        End If
    Else
        frmBank.lstInventory.AddItem i & "> Empty"
    End If
    DoEvents
Next i

For i = 1 To MAX_BANK
    If GetPlayerBankItemNum(MyIndex, i) > 0 Then
        If Item(GetPlayerBankItemNum(MyIndex, i)).Type = ITEM_TYPE_CURRENCY Then
            frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).name) & " (" & GetPlayerBankItemValue(MyIndex, i) & ")"
        Else
            If GetPlayerWeaponSlot(MyIndex) = i Or GetPlayerArmorSlot(MyIndex) = i Or GetPlayerHelmetSlot(MyIndex) = i Or GetPlayerShieldSlot(MyIndex) = i Or GetPlayerBootsSlot(MyIndex) = i Or GetPlayerGlovesSlot(MyIndex) = i Or GetPlayerRingSlot(MyIndex) = i Or GetPlayerAmuletSlot(MyIndex) = i Then
                frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).name) & " (worn)"
            Else
                frmBank.lstBank.AddItem i & "> " & Trim(Item(GetPlayerBankItemNum(MyIndex, i)).name)
            End If
        End If
    Else
        frmBank.lstBank.AddItem i & "> Empty"
    End If
    DoEvents
Next i
frmBank.lstBank.ListIndex = 0
frmBank.lstInventory.ListIndex = 0
End Sub
