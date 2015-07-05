Attribute VB_Name = "modRendering"

Option Explicit

' DirectX8 Object
Private Directx8 As Directx8 ' The master DirectX object.
Private Direct3D As Direct3D8 ' Controls all things 3D.
Public Direct3D_Device As Direct3DDevice8 ' Represents the hardware rendering.
Private Direct3DX As D3DX8

' The 2D (Transformed and Lit) vertex format.
Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE

Public Type GeomRec
    Top As Long
    Left As Long
    Height As Long
    Width As Long
End Type

' The 2D (Transformed and Lit) vertex format type.
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    RHW As Single
    Color As Long
    TU As Single
    TV As Single
End Type

Private Vertex_List(3) As TLVERTEX ' 4 vertices will make a square.

Public RenderingMode As Long

Private Direct3D_Window As D3DPRESENT_PARAMETERS ' Backbuffer and viewport description.
Private Display_Mode As D3DDISPLAYMODE

' Graphic Textures
Public Tex_Item() As DX8TextureRec ' Arrays
Public Tex_Character() As DX8TextureRec
Public Tex_Paperdoll() As DX8TextureRec
Public Tex_Tileset() As DX8TextureRec
Public Tex_Resource() As DX8TextureRec
Public Tex_Animation() As DX8TextureRec
Public Tex_SpellIcon() As DX8TextureRec
Public Tex_Face() As DX8TextureRec
Public Tex_Fog() As DX8TextureRec
Public Tex_Panorama() As DX8TextureRec
Public Tex_Emoticon() As DX8TextureRec
Public Tex_Blood As DX8TextureRec ' Singles
Public Tex_Misc As DX8TextureRec
Public Tex_Direction As DX8TextureRec
Public Tex_Target As DX8TextureRec
Public Tex_Bars As DX8TextureRec
Public Tex_Selection As DX8TextureRec
Public Tex_White As DX8TextureRec
Public Tex_Weather As DX8TextureRec
Public Tex_ChatBubble As DX8TextureRec
Public Tex_Fade As DX8TextureRec
Public Tex_Equip As DX8TextureRec
Public Tex_Base As DX8TextureRec

Public Quest_Start As DX8TextureRec
Public Quest_Start_U As DX8TextureRec
Public Quest_Finished As DX8TextureRec
Public Quest_Finished_U As DX8TextureRec
Public Quest_Progress As DX8TextureRec
Public Quest_Progress_U As DX8TextureRec
Public Quest_Completed As DX8TextureRec
Public Quest_Completed_U As DX8TextureRec

' Caching
Public lowerTilesCache As Direct3DTexture8
Public upperTilesCache As Direct3DTexture8
Public redrawMapCache As Boolean

' Character Editor Sprite
Public Tex_CharSprite As DX8TextureRec
Public LastCharSpriteTimer As Long
Public LastAdminSpriteTimer As Long
Private CharSpritePos As Byte
Private AdminSpritePos As Byte

' Number of graphic files
Public NumTileSets As Long
Public NumCharacters As Long
Public NumPaperdolls As Long
Public NumItems As Long
Public NumResources As Long
Public NumAnimations As Long
Public NumSpellIcons As Long
Public NumFaces As Long
Public NumFogs As Long
Public NumPanoramas As Long
Public NumEmoticons As Long

Public Type DX8TextureRec
    Texture As Long
    Width As Long
    Height As Long
    filepath As String
    TexWidth As Long
    TexHeight As Long
    ImageData() As Byte
    HasData As Boolean
End Type

Public Type GlobalTextureRec
    Texture As Direct3DTexture8
    TexWidth As Long
    TexHeight As Long
    Timer As Long
    Loaded As Boolean

    Static As Boolean
End Type

Public Type RECT
    Top As Long
    Left As Long
    Bottom As Long
    Right As Long
End Type

Public gTexture() As GlobalTextureRec
Public NumTextures As Long

' ********************
' ** Initialization **
' ********************
Public Function InitDX8() As Boolean
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set Directx8 = New Directx8 ' Creates the DirectX object.
    Set Direct3D = Directx8.Direct3DCreate() ' Creates the Direct3D object using the DirectX object.
    Set Direct3DX = New D3DX8
    
    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode ' Use the current display mode that you are on (resolution).
    Direct3D_Window.Windowed = True ' The app will be in windowed mode.
    
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_DISCARD ' Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format ' Sets the format that was retrieved into the backbuffer.
    
    ' Creates the rendering device with some useful info, along with the info
    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY
    Direct3D_Window.BackBufferCount = 1
    Display_Mode.Width = frmMain.picScreen.ScaleWidth
    Display_Mode.Height = frmMain.picScreen.ScaleHeight
    Direct3D_Window.BackBufferWidth = frmMain.picScreen.ScaleWidth
    Direct3D_Window.BackBufferHeight = frmMain.picScreen.ScaleHeight
    Direct3D_Window.hDeviceWindow = frmMain.picScreen.hWnd ' Use frmMain as the device window.
    
    ' We've already setup for Direct3D_Window.
    If TryCreateDirectX8Device = False Then
        MsgBox "Unable to initialize DirectX8. You may be missing dx8vb.dll or have incompatible hardware to use DirectX8."
        DestroyGame
    End If

    With Direct3D_Device
        .SetVertexShader FVF

        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False

        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE

        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0

        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    
    ' Initialize the surfaces
    LoadTextures
    
    ' We're done
    InitDX8 = True
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "InitDX8", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function Ceiling(dblValIn As Double, dblCeilIn As Double) As Double
    ' Round it
    Ceiling = Round(dblValIn / dblCeilIn, 0) * dblCeilIn
    
    ' If it rounded down, force it up
    If Ceiling < dblValIn Then Ceiling = Ceiling + dblCeilIn
End Function

Public Sub DestroyDX8()
    UnloadTextures
    
    Set Direct3DX = Nothing
    Set Direct3D_Device = Nothing
    Set Direct3D = Nothing
    Set Directx8 = Nothing
End Sub

Public Sub DrawGDI()
    ' Cycle through in-game stuff before cycling through editors
    If FormVisible("frmMenu") Then
        If frmMenu.picCharacter.Visible Then Menu_DrawCharacter
    End If
    
    If FormVisible("frmMain") Then
        If frmMain.picSpellDesc.Visible Then DrawSpellDesc LastSpellDesc
        If frmMain.picItemDesc.Visible Then DrawItemDesc LastItemDesc
        If frmMain.picHotbar.Visible Then DrawHotbar
        If frmMain.picInventory.Visible Then DrawInventory
        If frmMain.picCharacter.Visible Then DrawPlayerCharFace
        If frmMain.picEquipment.Visible Then DrawEquipment
        If frmMain.picChatFace.Visible Then DrawEventChatFace
        If frmMain.picSpells.Visible Then DrawPlayerSpells
        If frmMain.picShop.Visible Then DrawShop
        If frmMain.picTempBank.Visible Then DrawBankItem frmMain.picTempBank.Left, frmMain.picTempBank.Top
        If frmMain.picBank.Visible Then DrawBank
        If frmMain.picTrade.Visible Then DrawTrade
    End If
    
    If FormVisible("frmEditor_Animation") Then
        EditorAnim_DrawAnim
    End If
    
    If FormVisible("frmEditor_Item") Then
        EditorItem_DrawItem
        EditorItem_DrawPaperdoll
    End If
    
    If FormVisible("frmEditor_Map") Then
        If frmEditor_Map.picMapItem.Visible Then EditorMap_DrawMapItem
        
        ' Renders random tiles in map editor
        If frmEditor_Map.chkRandom.Value = 1 Then
            Call EditorMap_DrawRandom
        End If
    End If
    
    If FormVisible("frmMapPreview") Then
        If frmMapPreview.redrawMapPreview Then
            frmMapPreview.redrawMapPreview = False
            renderMapPreview
        End If
    End If
    
    ' Character editor
    If FormVisible("frmCharEditor") Then
        If Tex_CharSprite.Texture > 0 And requestedPlayer.Sprite > 0 Then
            If LastCharSpriteTimer + 300 < timeGetTime Then
                LastCharSpriteTimer = timeGetTime
                EditorChar_AnimSprite frmCharEditor.picSprite, frmCharEditor.txtSprite.text, CharSpritePos
            End If
        End If
    End If
    
    If FormVisible("frmEditor_MapProperties") Then
        EditorMapProperties_DrawPanorama
    End If
    
    If FormVisible("frmEditor_NPC") Then
        EditorNPC_DrawSprite
    End If
    
    If FormVisible("frmEditor_Resource") Then
        EditorResource_DrawSprite
    End If
    
    If FormVisible("frmEditor_Spell") Then
        EditorSpell_DrawIcon
    End If
    
    If FormVisible("frmEditor_Events") Then
        EditorEvent_DrawFace
        EditorEvent_DrawFace2
        EditorEvent_DrawGraphic
    End If
    
    If FormVisible("frmEditor_Emoticon") Then
        If frmEditor_Emoticon.Visible Then
            EditorEmoticon_DrawIcon
        End If
    End If
    
    If FormVisible("frmEditor_Class") Then
        With frmEditor_Class
            If .scrlMSprite.Visible Then
                Call EditorClass_DrawSprite(0)
            Else
                Call EditorClass_DrawSprite(1)
            End If
            
            If .scrlMFace.Visible Then
                Call EditorClass_DrawFace(0)
            Else
                Call EditorClass_DrawFace(1)
            End If
        End With
    End If
    If FormVisible("frmAdmin") Then
        If frmAdmin.txtSprite.text > 0 And LastAdminSpriteTimer + 300 < timeGetTime Then
            LastAdminSpriteTimer = timeGetTime
            EditorChar_AnimSprite frmAdmin.picSprite, frmAdmin.txtSprite.text, AdminSpritePos
        End If
        If ArrayIsInitialized(lastSpawnedItems) Then
            If UBound(lastSpawnedItems) > 0 Then
                drawRecentItem Item(lastSpawnedItems(frmAdmin.rcSwitcher.Value)).Pic
            End If
        End If
    End If
End Sub

Function TryCreateDirectX8Device() As Boolean
    Dim I As Long

    On Error GoTo nexti
    
    For I = 1 To 4
        Select Case I
            Case 1
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 2
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 3
                Set Direct3D_Device = Direct3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.picScreen.hWnd, D3DCREATE_MIXED_VERTEXPROCESSING, Direct3D_Window)
                TryCreateDirectX8Device = True
                Exit Function
            Case 4
                TryCreateDirectX8Device = False
                Exit Function
        End Select
nexti:
    Next

End Function

Function GetNearestPOT(Value As Long) As Long
    Dim I As Long

    Do While 2 ^ I < Value
        I = I + 1
    Loop
    
    GetNearestPOT = 2 ^ I
End Function

Public Sub SetTexture(ByRef TextureRec As DX8TextureRec)
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If TextureRec.HasData = False Then LoadTexture TextureRec
    Set gTexture(TextureRec.Texture).Texture = Direct3DX.CreateTextureFromFileInMemoryEx(Direct3D_Device, TextureRec.ImageData(0), UBound(TextureRec.ImageData) + 1, GetNearestPOT(TextureRec.Width), GetNearestPOT(TextureRec.Height), D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, ByVal (0), ByVal 0, ByVal 0)
    gTexture(TextureRec.Texture).TexWidth = GetNearestPOT(TextureRec.Width)
    gTexture(TextureRec.Texture).TexHeight = GetNearestPOT(TextureRec.Height)
    gTexture(TextureRec.Texture).Loaded = True
    gTexture(TextureRec.Texture).Timer = timeGetTime + 150000
    Exit Sub
   
    ' Error Handler
ErrorHandler:
    HandleError "SetTexture", "modGraphics", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub UnsetTexture(ByRef textureNum As Long, Optional ByVal Ignore As Boolean = False)
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If Ignore = True Then
        Set gTexture(textureNum).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(textureNum)), LenB(gTexture(textureNum))
        gTexture(textureNum).Timer = 0
        gTexture(textureNum).Loaded = False
    End If
    
    If gTexture(textureNum).Static = True Then
    
        Exit Sub    ' Don't unset static textures
    
    End If
    
    If gTexture(textureNum).Timer < timeGetTime And gTexture(textureNum).Timer <> 0 Then
        Set gTexture(textureNum).Texture = Nothing
        ZeroMemory ByVal VarPtr(gTexture(textureNum)), LenB(gTexture(textureNum))
        gTexture(textureNum).Timer = 0
        gTexture(textureNum).Loaded = False
    End If
    Exit Sub
   
' Error Handler
ErrorHandler:
    HandleError "UnsetTexture", "modGraphics", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub LoadTexture(ByRef TextureRec As DX8TextureRec)
    Dim SourceBitmap As cGDIpImage, ConvertedBitmap As cGDIpImage, GDIGraphics As cGDIpRenderer, GDIToken As cGDIpToken, I As Long
    Dim newWidth As Long, newHeight As Long, ImageData() As Byte, fn As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If TextureRec.HasData = False Then
        Set GDIToken = New cGDIpToken
        
        ' Make sure it loaded correctly
        If GDIToken.Token = 0& Then MsgBox "GDI+ failed to load, exiting game!": DestroyGame
        
        Set SourceBitmap = New cGDIpImage
        Call SourceBitmap.LoadPicture_FileName(TextureRec.filepath, GDIToken)
        TextureRec.Width = SourceBitmap.Width
        TextureRec.Height = SourceBitmap.Height
        
        newWidth = GetNearestPOT(TextureRec.Width)
        newHeight = GetNearestPOT(TextureRec.Height)
        
        If newWidth <> SourceBitmap.Width Or newHeight <> SourceBitmap.Height Then
            Set ConvertedBitmap = New cGDIpImage
            Set GDIGraphics = New cGDIpRenderer
            I = GDIGraphics.CreateGraphicsFromImageClass(SourceBitmap)
            Call ConvertedBitmap.LoadPicture_FromNothing(newWidth, newHeight, I, GDIToken) 'This is no longer backwards and it now works.
            Call GDIGraphics.DestroyHGraphics(I)
            I = GDIGraphics.CreateGraphicsFromImageClass(ConvertedBitmap)
            Call GDIGraphics.AttachTokenClass(GDIToken)
            Call GDIGraphics.RenderImageClassToHGraphics(SourceBitmap, I)
            Call ConvertedBitmap.SaveAsPNG(ImageData)
            GDIGraphics.DestroyHGraphics (I)
            TextureRec.ImageData = ImageData
            TextureRec.HasData = True
            Set ConvertedBitmap = Nothing
            Set GDIGraphics = Nothing
            Set SourceBitmap = Nothing
        Else
            Call SourceBitmap.SaveAsPNG(ImageData)
            TextureRec.ImageData = ImageData
            TextureRec.HasData = True
            Set SourceBitmap = Nothing
        End If
    Else
        ImageData = TextureRec.ImageData
        TextureRec.HasData = True
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "LoadTexture", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub LoadTextures()
    Dim I As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Call CheckTilesets
    Call CheckCharacters
    Call CheckPaperdolls
    Call CheckAnimations
    Call CheckItems
    Call CheckResources
    Call CheckSpellIcons
    Call CheckFaces
    Call CheckFogs
    Call CheckPanoramas
    Call CheckEmoticons
    
    NumTextures = NumTextures + 19
    
    ReDim Preserve gTexture(NumTextures)
    Quest_Completed.filepath = App.Path & "\data files\graphics\gui\main\quest_completed.png"
    Quest_Completed.Texture = NumTextures - 19
    LoadTexture Quest_Completed
    Quest_Completed_U.filepath = App.Path & "\data files\graphics\gui\main\quest_completed_u.png"
    Quest_Completed_U.Texture = NumTextures - 18
    LoadTexture Quest_Completed_U
    
    Quest_Start.filepath = App.Path & "\data files\graphics\gui\main\quest_start.png"
    Quest_Start.Texture = NumTextures - 17
    LoadTexture Quest_Start
    Quest_Start_U.filepath = App.Path & "\data files\graphics\gui\main\quest_start_u.png"
    Quest_Start_U.Texture = NumTextures - 16
    LoadTexture Quest_Start_U
    Quest_Finished.filepath = App.Path & "\data files\graphics\gui\main\quest_finished.png"
    Quest_Finished.Texture = NumTextures - 15
    LoadTexture Quest_Finished
    Quest_Finished_U.filepath = App.Path & "\data files\graphics\gui\main\quest_finished_u.png"
    Quest_Finished_U.Texture = NumTextures - 14
    LoadTexture Quest_Finished_U
    Quest_Progress.filepath = App.Path & "\data files\graphics\gui\main\quest_progress.png"
    Quest_Progress.Texture = NumTextures - 13
    LoadTexture Quest_Progress
    Quest_Progress_U.filepath = App.Path & "\data files\graphics\gui\main\quest_progress_u.png"
    Quest_Progress_U.Texture = NumTextures - 12
    LoadTexture Quest_Progress_U
    
    Tex_Base.filepath = App.Path & "\data files\graphics\gui\main\base.png"
    Tex_Base.Texture = NumTextures - 11
    LoadTexture Tex_Base
    Tex_Equip.filepath = App.Path & "\data files\graphics\gui\main\equip.png"
    Tex_Equip.Texture = NumTextures - 10
    LoadTexture Tex_Equip
    Tex_Fade.filepath = App.Path & "\data files\graphics\misc\fader.png"
    Tex_Fade.Texture = NumTextures - 9
    LoadTexture Tex_Fade
    Tex_ChatBubble.filepath = App.Path & "\data files\graphics\misc\chatbubble.png"
    Tex_ChatBubble.Texture = NumTextures - 8
    LoadTexture Tex_ChatBubble
    Tex_Weather.filepath = App.Path & "\data files\graphics\misc\weather.png"
    Tex_Weather.Texture = NumTextures - 7
    LoadTexture Tex_Weather
    Tex_White.filepath = App.Path & "\data files\graphics\misc\white.png"
    Tex_White.Texture = NumTextures - 6
    LoadTexture Tex_White
    Tex_Direction.filepath = App.Path & "\data files\graphics\misc\direction.png"
    Tex_Direction.Texture = NumTextures - 5
    LoadTexture Tex_Direction
    Tex_Target.filepath = App.Path & "\data files\graphics\misc\target.png"
    Tex_Target.Texture = NumTextures - 4
    LoadTexture Tex_Target
    Tex_Misc.filepath = App.Path & "\data files\graphics\misc\misc.png"
    Tex_Misc.Texture = NumTextures - 3
    LoadTexture Tex_Misc
    Tex_Blood.filepath = App.Path & "\data files\graphics\misc\blood.png"
    Tex_Blood.Texture = NumTextures - 2
    LoadTexture Tex_Blood
    Tex_Bars.filepath = App.Path & "\data files\graphics\misc\bars.png"
    Tex_Bars.Texture = NumTextures - 1
    LoadTexture Tex_Bars
    Tex_Selection.filepath = App.Path & "\data files\graphics\misc\select.png"
    Tex_Selection.Texture = NumTextures
    LoadTexture Tex_Selection
    
    EngineInitFontTextures
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "LoadTextures", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
Public Function ArrayIsInitialized2(arr) As Boolean

  Dim memVal As Long

  CopyMemory memVal, ByVal VarPtr(arr) + 8, ByVal 4 'get pointer to array
  CopyMemory memVal, ByVal memVal, ByVal 4  'see if it points to an address...
  ArrayIsInitialized2 = (memVal <> 0)        '...if it does, array is intialized

End Function
Public Sub UnloadTextures()
Dim I As Long
    
    ' If debug mode, handle error then exit out
    On Error Resume Next
    
        For I = 1 To NumTextures
            Set gTexture(I).Texture = Nothing
            ZeroMemory ByVal VarPtr(gTexture(I)), LenB(gTexture(I))
        Next
        
        ReDim gTexture(1)

    For I = 1 To NumTileSets
        Tex_Tileset(I).Texture = 0
    Next

    For I = 1 To NumItems
        Tex_Item(I).Texture = 0
    Next

    For I = 1 To NumCharacters
        Tex_Character(I).Texture = 0
    Next
    
    For I = 1 To NumPaperdolls
        Tex_Paperdoll(I).Texture = 0
    Next
    
    For I = 1 To NumResources
        Tex_Resource(I).Texture = 0
    Next
    
    For I = 1 To NumAnimations
        Tex_Animation(I).Texture = 0
    Next
    
    For I = 1 To NumSpellIcons
        Tex_SpellIcon(I).Texture = 0
    Next
    
    For I = 1 To NumFaces
        Tex_Face(I).Texture = 0
    Next
    
    For I = 1 To NumPanoramas
        Tex_Panorama(I).Texture = 0
    Next
    
    For I = 1 To NumEmoticons
        Tex_Emoticon(I).Texture = 0
    Next

    Tex_Equip.Texture = 0
    Tex_Base.Texture = 0
    Tex_Fade.Texture = 0
    Tex_ChatBubble.Texture = 0
    Tex_Weather.Texture = 0
    Tex_White.Texture = 0
    Tex_Bars.Texture = 0
    Tex_Misc.Texture = 0
    Tex_Blood.Texture = 0
    Tex_Direction.Texture = 0
    Tex_Target.Texture = 0
    Tex_Selection.Texture = 0
    Quest_Start.Texture = 0
    Quest_Start_U.Texture = 0
    Quest_Finished.Texture = 0
    Quest_Finished_U.Texture = 0
    Quest_Progress.Texture = 0
    Quest_Progress_U.Texture = 0
    
    UnloadFontTextures
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "UnloadTextures", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' **************
' ** Drawing **
' **************
Public Sub renderMapPreview()
    Dim destRECT As D3DRECT

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderCache lowerTilesCache, 0, 0, 0, 0, frmMapPreview.picMapPreview.Width, frmMapPreview.picMapPreview.Height, Map.MaxX * PIC_X, Map.MaxY * PIC_Y
    RenderCache upperTilesCache, 0, 0, 0, 0, frmMapPreview.picMapPreview.Width, frmMapPreview.picMapPreview.Height, Map.MaxX * PIC_X, Map.MaxY * PIC_Y
    
    With destRECT
        .X1 = 0
        .X2 = frmMapPreview.picMapPreview.Width
        .Y1 = 0
        .Y2 = frmMapPreview.picMapPreview.Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmMapPreview.picMapPreview.hWnd, ByVal (0)

    Exit Sub
    
    ' Error handler
ErrorHandler:
    HandleError "renderMapPreview", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub RenderTexture(ByRef TextureRec As DX8TextureRec, ByVal dX As Single, ByVal dY As Single, ByVal sX As Single, ByVal sY As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional Color As Long = -1)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Dim textureNum As Long
    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    
    textureNum = TextureRec.Texture
    
    If gTexture(textureNum).Loaded = False Then
        SetTexture TextureRec
    Else
        gTexture(textureNum).Timer = timeGetTime + 150000
    End If
    
    textureWidth = gTexture(textureNum).TexWidth
    textureHeight = gTexture(textureNum).TexHeight
    
    If sY + sHeight > textureHeight Then Exit Sub
    If sX + sWidth > textureWidth Then Exit Sub
    If sX < 0 Then Exit Sub
    If sY < 0 Then Exit Sub

    sX = sX - 0.5
    sY = sY - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sX / textureWidth)
    sourceY = (sY / textureHeight)
    sourceWidth = ((sX + sWidth) / textureWidth)
    sourceHeight = ((sY + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, Color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, Color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, Color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, Color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    Direct3D_Device.SetTexture 0, gTexture(textureNum).Texture
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "RenderTexture", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub RenderCache(ByRef mapChache As Direct3DTexture8, ByVal dX As Single, ByVal dY As Single, ByVal sX As Single, ByVal sY As Single, ByVal dWidth As Single, ByVal dHeight As Single, ByVal sWidth As Single, ByVal sHeight As Single, Optional Color As Long = -1)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Dim textureWidth As Long, textureHeight As Long, sourceX As Single, sourceY As Single, sourceWidth As Single, sourceHeight As Single
    
    'Setting up Texture Source
    Dim infoS As D3DSURFACE_DESC
    Dim textSurf As Direct3DSurface8
    
    Set textSurf = mapChache.GetSurfaceLevel(0)
    textSurf.GetDesc infoS
    
    textureWidth = infoS.Width
    textureHeight = infoS.Height
    
    If sY + sHeight > textureHeight Then Exit Sub
    If sX + sWidth > textureWidth Then Exit Sub
    If sX < 0 Then Exit Sub
    If sY < 0 Then Exit Sub

    sX = sX - 0.5
    sY = sY - 0.5
    dY = dY - 0.5
    dX = dX - 0.5
    sWidth = sWidth
    sHeight = sHeight
    dWidth = dWidth
    dHeight = dHeight
    sourceX = (sX / textureWidth)
    sourceY = (sY / textureHeight)
    sourceWidth = ((sX + sWidth) / textureWidth)
    sourceHeight = ((sY + sHeight) / textureHeight)
    
    Vertex_List(0) = Create_TLVertex(dX, dY, 0, 1, Color, 0, sourceX + 0.000003, sourceY + 0.000003)
    Vertex_List(1) = Create_TLVertex(dX + dWidth, dY, 0, 1, Color, 0, sourceWidth + 0.000003, sourceY + 0.000003)
    Vertex_List(2) = Create_TLVertex(dX, dY + dHeight, 0, 1, Color, 0, sourceX + 0.000003, sourceHeight + 0.000003)
    Vertex_List(3) = Create_TLVertex(dX + dWidth, dY + dHeight, 0, 1, Color, 0, sourceWidth + 0.000003, sourceHeight + 0.000003)
    
    Direct3D_Device.SetTexture 0, mapChache
    Direct3D_Device.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex_List(0), Len(Vertex_List(0))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "RenderTexture", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub RenderTextureByRects(TextureRec As DX8TextureRec, sRECT As RECT, dRect As RECT)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    RenderTexture TextureRec, dRect.Left, dRect.Top, sRECT.Left, sRECT.Top, dRect.Right - dRect.Left, dRect.Bottom - dRect.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "RenderTextureByRects", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' This function will make it much easier to setup the vertices with the info it needs.
Private Function Create_TLVertex(X As Single, Y As Single, Z As Single, RHW As Single, Color As Long, Specular As Long, TU As Single, TV As Single) As TLVERTEX
    Create_TLVertex.X = X
    Create_TLVertex.Y = Y
    Create_TLVertex.Z = Z
    Create_TLVertex.RHW = RHW
    Create_TLVertex.Color = Color
    Create_TLVertex.TU = TU
    Create_TLVertex.TV = TV
End Function

Public Sub DrawGrid(ByVal X As Long, ByVal Y As Long)
    Dim Top As Long, Left As Long
    
    ' Render grid
    Top = 24
    Left = 0

    RenderTexture Tex_Direction, ConvertMapX(X * PIC_X), ConvertMapY(Y * PIC_Y), Left, Top, 32, 32, 32, 32
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal X As Long, ByVal Y As Long)
    Dim I As Long, Top As Long, Left As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Render dir blobs
    For I = 1 To 4
        Left = (I - 1) * 8
        
        ' Find out whether render blocked or not
        If Not IsDirBlocked(Map.Tile(X, Y).DirBlock, CByte(I)) Then
            Top = 8
        Else
            Top = 16
        End If
       
        RenderTexture Tex_Direction, ConvertMapX(X * PIC_X) + DirArrowX(I), ConvertMapY(Y * PIC_Y) + DirArrowY(I), Left, Top, 8, 8, 8, 8
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawDirection", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawTarget(ByVal X As Long, ByVal Y As Long)
    Dim sRECT As RECT
    Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = 0
        .Right = Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' Clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Target, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawTarget", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawHover(ByVal tType As Long, ByVal Target As Long, ByVal X As Long, ByVal Y As Long)
    Dim sRECT As RECT
    Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Tex_Target.Texture = 0 Then Exit Sub
    
    Width = Tex_Target.Width / 2
    Height = Tex_Target.Height

    With sRECT
        .Top = 0
        .Bottom = Height
        .Left = Width
        .Right = .Left + Width
    End With
    
    X = X - ((Width - 32) / 2)
    Y = Y - (Height / 2)

    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    ' Clipping
    If Y < 0 Then
        With sRECT
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With sRECT
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Target, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawHover", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawWholeMapLowerTiles(ByVal X As Long, ByVal Y As Long)
    Dim rec As RECT
    Dim I As Long, Alpha As Byte
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    With Map.Tile(X, Y)
        For I = MapLayer.Ground To MapLayer.Cover
            If InMapEditor And I < CurrentLayer Then
                If frmMain.chkDimLayers.Value = 1 Then
                    Alpha = 255 - ((CurrentLayer - I) * 48)
                Else
                    Alpha = 255
                End If
            Else
                Alpha = 255
            End If
            
            If Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(I).Tileset), X * PIC_X, Y * PIC_Y, .Layer(I).X * 32, .Layer(I).Y * 32, 32, 32, 32, 32, D3DColorARGB(Alpha, 255, 255, 255)
            ElseIf Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_AUTOTILE And Options.Autotile = 1 Then
                ' Draw autotiles
                DrawAutoTile I, X * PIC_X, Y * PIC_Y, 1, X, Y, Alpha
                DrawAutoTile I, (X * PIC_X) + 16, Y * PIC_Y, 2, X, Y, Alpha
                DrawAutoTile I, X * PIC_X, (Y * PIC_Y) + 16, 3, X, Y, Alpha
                DrawAutoTile I, (X * PIC_X) + 16, (Y * PIC_Y) + 16, 4, X, Y, Alpha
            End If
        Next
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawWholeMapLowerTiles", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawWholeMapUpperTiles(ByVal X As Long, ByVal Y As Long)
    Dim rec As RECT
    Dim I As Long, Alpha As Byte

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    With Map.Tile(X, Y)
        For I = MapLayer.Fringe To MapLayer.Roof
            If I < CurrentLayer And InMapEditor Then
                If frmMain.chkDimLayers.Value = 1 Then ' has to be here cause checking for it in previous IF would load it to memory
                    Alpha = 255 - ((CurrentLayer - I) * 48)
                Else
                    Alpha = 255
                End If
            Else
                Alpha = 255
            End If
            
            If Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(I).Tileset), X * PIC_X, Y * PIC_Y, .Layer(I).X * 32, .Layer(I).Y * 32, 32, 32, 32, 32, D3DColorARGB(Alpha, 255, 255, 255)
            ElseIf Autotile(X, Y).Layer(I).RenderState = RENDER_STATE_AUTOTILE And Options.Autotile = 1 Then
                ' Draw autotiles
                DrawAutoTile I, X * PIC_X, Y * PIC_Y, 1, X, Y, Alpha
                DrawAutoTile I, (X * PIC_X) + 16, Y * PIC_Y, 2, X, Y, Alpha
                DrawAutoTile I, (X * PIC_X), (Y * PIC_Y) + 16, 3, X, Y, Alpha
                DrawAutoTile I, (X * PIC_X) + 16, (Y * PIC_Y) + 16, 4, X, Y, Alpha
            End If
        Next
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawWholeMapUpperTiles", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawBlood(ByVal Index As Long)
    Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Load blood
    BloodCount = Tex_Blood.Width / 32
    
    With Blood(Index)
        If .Alpha <= 0 Then
            Call ClearBlood(Index)
            Exit Sub
        End If
        
        ' Check if we should be seeing it
        If .Timer + 20000 < timeGetTime Then
            .Alpha = .Alpha - 1
        End If
        
        rec.Top = 0
        rec.Bottom = PIC_Y
        rec.Left = (.Sprite - 1) * PIC_X
        rec.Right = rec.Left + PIC_X
        RenderTexture Tex_Blood, ConvertMapX(.X * PIC_X), ConvertMapY(.Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorARGB(Blood(Index).Alpha, 255, 255, 255)
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawBlood", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
    Dim Sprite As Integer, sRECT As GeomRec, I As Long, Width As Long, Height As Long, looptime As Long, FrameCount As Long
    Dim X As Long, Y As Long, LockIndex As Long
    
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    Sprite = Animation(AnimInstance(Index).Animation).Sprite(Layer)
    
    If Sprite < 1 Or Sprite > NumAnimations Then Exit Sub
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' Total width divided by frame count
    Width = Tex_Animation(Sprite).Width \ AnimColumns
    Height = Width

    With sRECT
        .Top = (Height * ((AnimInstance(Index).frameIndex(Layer) - 1) \ AnimColumns))
        .Height = Height
        .Left = (Width * (((AnimInstance(Index).frameIndex(Layer) - 1) Mod AnimColumns)))
        .Width = Width
    End With
    
    ' Change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' If <> none
        ' Is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' Quick save the index
            LockIndex = AnimInstance(Index).LockIndex
            ' Check if is ingame
            If IsPlaying(LockIndex) Then
                ' Check if on same map
                If GetPlayerMap(LockIndex) = GetPlayerMap(MyIndex) Then
                    ' Is on map, is playing, set x & y
                    X = (GetPlayerX(LockIndex) * PIC_X) + 16 - (Width / 2) + TempPlayer(LockIndex).xOffset
                    Y = (GetPlayerY(LockIndex) * PIC_Y) + 16 - (Height / 2) + TempPlayer(LockIndex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' Quick save the index
            LockIndex = AnimInstance(Index).LockIndex
            ' Check if NPC exists
            If MapNPC(LockIndex).num > 0 Then
                ' Check if alive
                If MapNPC(LockIndex).Vital(Vitals.HP) > 0 Then
                    ' Exists, is alive, set x & y
                    X = (MapNPC(LockIndex).X * PIC_X) + 16 - (Width / 2) + MapNPC(LockIndex).xOffset
                    Y = (MapNPC(LockIndex).Y * PIC_Y) + 16 - (Height / 2) + MapNPC(LockIndex).yOffset
                Else
                    ' NPC not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' NPC not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' No lock, default x + y
        X = (AnimInstance(Index).X * 32) + 16 - (Width / 2)
        Y = (AnimInstance(Index).Y * 32) + 16 - (Height / 2)
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y)
    
    RenderTexture Tex_Animation(Sprite), X, Y, sRECT.Left, sRECT.Top, sRECT.Width, sRECT.Height, sRECT.Width, sRECT.Height
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "DrawAnimation", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawMapItem(ByVal ItemNum As Long)
    Dim picNum As Integer, X As Long, I As Long
    Dim rec As RECT
    Dim MaxFrames As Byte

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    X = 0
    
    ' If it's not ours then don't render
    If X = 0 Then
        If Not Trim$(MapItem(ItemNum).PlayerName) = GetPlayerName(MyIndex) Then
            If Not Trim$(MapItem(ItemNum).PlayerName) = vbNullString Then Exit Sub
        End If
    End If

    picNum = Item(MapItem(ItemNum).num).Pic

    If picNum < 1 Or picNum > NumItems Then Exit Sub
    
    If Tex_Item(picNum).Width > PIC_X Then ' Has more than 1 frame
        With rec
            .Top = 0
            .Bottom = PIC_Y
            .Left = (MapItem(ItemNum).Frame * PIC_X)
            .Right = .Left + PIC_X
        End With
    Else
        With rec
            .Top = 0
            .Bottom = PIC_Y
            .Left = 0
            .Right = PIC_X
        End With
    End If
    
    RenderTexture Tex_Item(picNum), ConvertMapX(MapItem(ItemNum).X * PIC_X), ConvertMapY(MapItem(ItemNum).Y * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawMapItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawMapResource(ByVal Resource_num As Long)
    Dim Resource_Master As Long
    Dim Resource_State As Long
    Dim Resource_Sprite As Long
    Dim rec As RECT
    Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Make sure it's not out of map
    If MapResource(Resource_num).X > Map.MaxX Then Exit Sub
    If MapResource(Resource_num).Y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_Master = Map.Tile(MapResource(Resource_num).X, MapResource(Resource_num).Y).Data1
    
    If Resource_Master = 0 Then Exit Sub

    If Resource(Resource_Master).ResourceImage = 0 Then Exit Sub
    
    ' Get the Resource state
    Resource_State = MapResource(Resource_num).ResourceState

    If Resource_State = 0 Then ' Normal
        Resource_Sprite = Resource(Resource_Master).ResourceImage
    ElseIf Resource_State = 1 Then ' Used
        Resource_Sprite = Resource(Resource_Master).ExhaustedImage
    End If
    
    ' Cut down everything if we're editing
    If InMapEditor Then
        Resource_Sprite = Resource(Resource_Master).ExhaustedImage
    End If

    ' Src rect
    With rec
        .Top = 0
        .Bottom = Tex_Resource(Resource_Sprite).Height
        .Left = 0
        .Right = Tex_Resource(Resource_Sprite).Width
    End With

    ' Set base x + y, then the offset due to size
    X = (MapResource(Resource_num).X * PIC_X) - (Tex_Resource(Resource_Sprite).Width / 2) + 16
    Y = (MapResource(Resource_num).Y * PIC_Y) - Tex_Resource(Resource_Sprite).Height + 32
    
    ' Render it
    Call DrawResource(Resource_Sprite, X, Y, rec)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawMapResource", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub DrawResource(ByVal Resource As Long, ByVal dX As Long, dY As Long, rec As RECT)
    Dim X As Long
    Dim Y As Long
    Dim Width As Long
    Dim Height As Long
    Dim destRECT As RECT

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Resource < 1 Or Resource > NumResources Then Exit Sub

    X = ConvertMapX(dX)
    Y = ConvertMapY(dY)
    
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    RenderTexture Tex_Resource(Resource), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawResource", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub DrawBars()
    Dim tmpy As Long, tmpx As Long
    Dim sWidth As Long, sHeight As Long
    Dim sRECT As RECT
    Dim I As Long, npcNum As Long, PartyIndex As Long, BarWidth As Long, MoveSpeed As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Dynamic bar calculations
    sWidth = Tex_Bars.Width
    sHeight = Tex_Bars.Height / 4
    
    ' Render health bars and casting bar
    For I = 1 To MAX_MAP_NPCS
        npcNum = MapNPC(I).num
        ' Exists
        If npcNum > 0 Then
            If Options.NPCVitals = 1 Then
                ' Alive
                If MapNPC(I).Vital(Vitals.HP) < NPC(npcNum).HP Then
                    ' lock to npc
                    tmpx = MapNPC(I).X * PIC_X + MapNPC(I).xOffset + 16 - (sWidth / 2)
                    tmpy = MapNPC(I).Y * PIC_Y + MapNPC(I).yOffset + 35
                    
                    ' Calculate the width to fill
                    BarWidth = ((MapNPC(I).Vital(Vitals.HP) / sWidth) / (NPC(npcNum).HP / sWidth)) * sWidth
                    
                    ' Draw bar background
                    With sRECT
                        .Top = sHeight * 3 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' Draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + BarWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
        
                If MapNPC(I).Vital(Vitals.MP) < NPC(npcNum).MP Then
                    ' lock to npc
                    tmpx = MapNPC(I).X * PIC_X + MapNPC(I).xOffset + 16 - (sWidth / 2)
                    
                    If MapNPC(I).Vital(Vitals.HP) = NPC(npcNum).HP Then
                        tmpy = MapNPC(I).Y * PIC_Y + MapNPC(I).yOffset + 35
                    Else
                        tmpy = MapNPC(I).Y * PIC_Y + MapNPC(I).yOffset + 35 + sHeight
                    End If
                    
                    ' Calculate the width to fill
                    BarWidth = ((MapNPC(I).Vital(Vitals.MP) / sWidth) / (NPC(npcNum).MP / sWidth)) * sWidth
                    
                    ' Draw bar background
                    With sRECT
                        .Top = sHeight * 3 ' MP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' Draw the bar proper
                    With sRECT
                        .Top = sHeight * 1 ' MP bar
                        .Left = 0
                        .Right = .Left + BarWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
            
            ' Check for npc casting time bar
            If MapNPC(I).SpellBuffer > 0 Then
                If MapNPC(I).SpellBufferTimer > timeGetTime - (Spell(MapNPC(I).SpellBuffer).CastTime * 1000) Then
                    ' lock to player
                    tmpx = MapNPC(I).X * PIC_X + MapNPC(I).xOffset + 16 - (sWidth / 2)

                    If Options.NPCVitals = 0 Or (MapNPC(I).Vital(Vitals.HP) = NPC(npcNum).HP And MapNPC(I).Vital(Vitals.MP) = NPC(npcNum).MP) Then
                        tmpy = MapNPC(I).Y * PIC_Y + MapNPC(I).yOffset + 35
                    Else
                        tmpy = MapNPC(I).Y * PIC_Y + MapNPC(I).yOffset + 35 + sHeight
                    End If
                   
                    ' Calculate the width to fill
                    BarWidth = (timeGetTime - MapNPC(I).SpellBufferTimer) / ((Spell(MapNPC(I).SpellBuffer).CastTime * 1000)) * sWidth

                    ' Draw bar background
                    With sRECT
                        .Top = sHeight * 3 ' Cooldown bar background
                        .Left = 0
                        .Right = sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' Draw the bar proper
                    With sRECT
                        .Top = sHeight * 2 ' Cooldown bar
                        .Left = 0
                        .Right = BarWidth
                        .Bottom = .Top + sHeight
                        
                    End With
                    RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
        End If
    Next
    
    If Options.PlayerVitals = 1 Then
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If GetPlayerMap(MyIndex) = GetPlayerMap(I) Then
                    ' Draw own health bar
                    If GetPlayerVital(I, Vitals.HP) < GetPlayerMaxVital(I, Vitals.HP) Then
                        ' lock to Player
                        tmpx = GetPlayerX(I) * PIC_X + TempPlayer(I).xOffset + 16 - (sWidth / 2)
                        tmpy = GetPlayerY(I) * PIC_X + TempPlayer(I).yOffset + 35
                    
                        ' Calculate the width to fill
                        BarWidth = ((GetPlayerVital(I, Vitals.HP) / sWidth) / (GetPlayerMaxVital(I, Vitals.HP) / sWidth)) * sWidth
                        
                        ' Draw bar background
                        With sRECT
                            .Top = sHeight * 3 ' HP bar background
                            .Left = 0
                            .Right = .Left + sWidth
                            .Bottom = .Top + sHeight
                        End With
                        
                        RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                       
                        ' Draw the bar proper
                        With sRECT
                            .Top = 0 ' HP bar
                            .Left = 0
                            .Right = .Left + BarWidth
                            .Bottom = .Top + sHeight
                        End With
                        
                        RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    End If
                    
                    ' Draw own mana bar
                    If GetPlayerVital(I, Vitals.MP) < GetPlayerMaxVital(I, Vitals.MP) Then
                        ' lock to Player
                        tmpx = GetPlayerX(I) * PIC_X + TempPlayer(I).xOffset + 16 - (sWidth / 2)
                        
                        If GetPlayerVital(I, HP) = GetPlayerMaxVital(I, Vitals.HP) Then
                            tmpy = GetPlayerY(I) * PIC_Y + TempPlayer(I).yOffset + 35
                        Else
                            tmpy = GetPlayerY(I) * PIC_Y + TempPlayer(I).yOffset + 35 + sHeight
                        End If
                       
                        ' Calculate the width to fill
                        BarWidth = ((GetPlayerVital(I, Vitals.MP) / sWidth) / (GetPlayerMaxVital(I, Vitals.MP) / sWidth)) * sWidth
                       
                        ' Draw bar background
                        With sRECT
                            .Top = sHeight * 3 ' MP bar background
                            .Left = 0
                            .Right = .Left + sWidth
                            .Bottom = .Top + sHeight
                        End With
                        
                        RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                       
                        ' Draw the bar proper
                        With sRECT
                            .Top = sHeight * 1 ' MP bar
                            .Left = 0
                            .Right = .Left + BarWidth
                            .Bottom = .Top + sHeight
                        End With
                        
                        RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    End If
                End If
            End If
        Next
    End If
                
    ' Check for player casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer)).CastTime > 0 Then
            ' lock to player
            tmpx = GetPlayerX(MyIndex) * PIC_X + TempPlayer(MyIndex).xOffset + 16 - (sWidth / 2)
            
            If Options.PlayerVitals = 0 Or (GetPlayerVital(I, HP) = GetPlayerMaxVital(I, Vitals.HP) And GetPlayerVital(I, MP) = GetPlayerMaxVital(I, MP)) Then
                tmpy = GetPlayerY(MyIndex) * PIC_Y + TempPlayer(MyIndex).yOffset + 35
            Else
                tmpy = GetPlayerY(MyIndex) * PIC_Y + TempPlayer(MyIndex).yOffset + 35 + sHeight
            End If
            
            ' Calculate the width to fill
            BarWidth = (timeGetTime - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer)).CastTime * 1000)) * sWidth
            
            ' Draw bar background
            With sRECT
                .Top = sHeight * 3 ' Cooldown bar background
                .Left = 0
                .Right = sWidth
                .Bottom = .Top + sHeight
            End With
            
            RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            
            ' Draw the bar proper
            With sRECT
                .Top = sHeight * 2 ' Cooldown bar
                .Left = 0
                .Right = BarWidth
                .Bottom = .Top + sHeight
            End With
            
            RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
        End If
    End If
    
    ' Draw party health bars
    If Party.num > 0 Then
        For I = 1 To MAX_PARTY_MEMBERS
            PartyIndex = Party.Member(I)
            If (PartyIndex > 0) And (Not PartyIndex = MyIndex) And (GetPlayerMap(PartyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(PartyIndex, Vitals.HP) > 0 And GetPlayerVital(PartyIndex, Vitals.HP) < GetPlayerMaxVital(PartyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpx = GetPlayerX(PartyIndex) * PIC_X + TempPlayer(PartyIndex).xOffset + 16 - (sWidth / 2)
                    tmpy = GetPlayerY(PartyIndex) * PIC_X + TempPlayer(PartyIndex).yOffset + 35
                    
                    ' Calculate the width to fill
                    BarWidth = ((GetPlayerVital(PartyIndex, Vitals.HP) / sWidth) / (GetPlayerMaxVital(PartyIndex, Vitals.HP) / sWidth)) * sWidth
                    
                    ' Draw bar background
                    With sRECT
                        .Top = sHeight * 1 ' HP bar background
                        .Left = 0
                        .Right = .Left + sWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    ' Draw the bar proper
                    With sRECT
                        .Top = 0 ' HP bar
                        .Left = 0
                        .Right = .Left + BarWidth
                        .Bottom = .Top + sHeight
                    End With
                    
                    RenderTexture Tex_Bars, ConvertMapX(tmpx), ConvertMapY(tmpy), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            End If
        Next
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawBars", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawHotbar()
    Dim sRECT As RECT, dRect As RECT, I As Long, num As String, n As Long, destRECT As D3DRECT

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To MAX_HOTBAR
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
        Direct3D_Device.BeginScene
    
        With dRect
            .Top = HotbarTop
            .Left = HotbarLeft + ((HotbarOffsetX + 32) * (((I - 1) Mod MAX_HOTBAR)))
            .Bottom = .Top + 32
            .Right = .Left + 32
        End With
        
        With destRECT
            .Y1 = HotbarTop
            .X1 = HotbarLeft + ((HotbarOffsetX + 32) * (((I - 1) Mod MAX_HOTBAR)))
            .Y2 = .Y1 + 32
            .X2 = .X1 + 32
        End With
        
        With sRECT
            .Top = 0
            .Left = 0
            .Bottom = 32
            .Right = 32
        End With
    
        Select Case Hotbar(I).sType
            Case 1 ' Inventory
                If Len(Item(Hotbar(I).Slot).Name) > 0 Then
                    If Item(Hotbar(I).Slot).Pic > 0 Then
                        If Item(Hotbar(I).Slot).Pic <= NumItems Then
                            RenderTextureByRects Tex_Item(Item(Hotbar(I).Slot).Pic), sRECT, dRect
                        End If
                    End If
                End If
            Case 2 ' Spell
                If Len(Spell(Hotbar(I).Slot).Name) > 0 Then
                    If Spell(Hotbar(I).Slot).Icon > 0 Then
                        With sRECT
                            .Top = 0
                            .Left = 0
                            .Bottom = 32
                            .Right = 32
                        End With
                        
                        ' Check for cooldown
                        For n = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(n) = Hotbar(I).Slot Then
                                ' has spell
                                If Not SpellCD(n) = 0 Then
                                    sRECT.Left = 32
                                    sRECT.Right = 64
                                    Exit For
                                End If
                            End If
                        Next
                        RenderTextureByRects Tex_SpellIcon(Spell(Hotbar(I).Slot).Icon), sRECT, dRect
                    End If
                End If
        End Select
    
        ' Render the letters
        If Options.WASD = 1 Then
            If I = 10 Then
                num = " 0"
            ElseIf I = 11 Then
                num = " -"
            ElseIf I = 12 Then
                num = " +"
            Else
                num = " " & Trim$(I)
            End If
        Else
            num = " F" & Trim$(I)
        End If
        RenderText Font_Default, num, dRect.Left + 2, dRect.Top + 16, White
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRECT, destRECT, frmMain.picHotbar.hWnd, ByVal (0)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawHotbar", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawPlayer(ByVal Index As Long)
    Dim Anim As Byte, I As Long, X As Long, Y As Long
    Dim Sprite As Long, spritetop As Long
    Dim rec As RECT
    Dim AttackSpeed As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Sprite = GetPlayerSprite(Index)

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    ' Speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        AttackSpeed = Item(GetPlayerEquipment(Index, Weapon)).WeaponSpeed
    Else
        AttackSpeed = 1000
    End If
    
    If AttackSpeed < 500 Then AttackSpeed = 500

    ' Reset frame
    If TempPlayer(Index).Moving > 0 Then
        Anim = TempPlayer(Index).Step
    Else
        Anim = 0
    End If
    
    ' If the sprite is constantly animated, make it animate
    If Class(GetPlayerClass(Index)).Animated = 0 Or TempPlayer(Index).Moving > 0 Then
        ' Check for attacking animation
        If TempPlayer(Index).AttackTimer + (AttackSpeed / 2) > timeGetTime Then
            If TempPlayer(Index).Attacking = 1 Then
                Anim = 3
            End If
        End If
    
        ' Check to see if we want to stop making him attack
        With TempPlayer(Index)
            If .AttackTimer + AttackSpeed < timeGetTime Then
                .Attacking = 0
                .AttackTimer = 0
            End If
        End With
    Else
        If TempPlayer(Index).AnimTimer + 100 <= timeGetTime Then
            TempPlayer(Index).Anim = TempPlayer(Index).Anim + 1
            If TempPlayer(Index).Anim >= 4 Then TempPlayer(Index).Anim = 0
            TempPlayer(Index).AnimTimer = timeGetTime
        End If
        Anim = TempPlayer(Index).Anim
    End If

    ' Set the left
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
        Case DIR_UPLEFT
            spritetop = 1
        Case DIR_UPRIGHT
            spritetop = 2
        Case DIR_DOWNLEFT
            spritetop = 1
        Case DIR_DOWNRIGHT
            spritetop = 2
    End Select

    With rec
        .Top = spritetop * (Tex_Character(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Character(Sprite).Height / 4)
        .Left = Anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    X = GetPlayerX(Index) * PIC_X + TempPlayer(Index).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32?
    If (Tex_Character(Sprite).Height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = GetPlayerY(Index) * PIC_Y + TempPlayer(Index).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = GetPlayerY(Index) * PIC_Y + TempPlayer(Index).yOffset
    End If

    ' Render the actual sprite
    Call DrawSprite(Sprite, X, Y, rec)
    
    ' Check for paperdolling
    Dim Size As Long
    Size = UBound(PaperdollOrder)
    For I = 1 To Size
        If GetPlayerEquipment(Index, PaperdollOrder(I)) > 0 Then
            If Item(GetPlayerEquipment(Index, PaperdollOrder(I))).Paperdoll > 0 Then
                Call DrawPaperdoll(X, Y, Item(GetPlayerEquipment(Index, PaperdollOrder(I))).Paperdoll, Anim, spritetop)
            End If
        End If
    Next
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "DrawPlayer", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawNPC(ByVal MapNPCNum As Long)
    Dim Anim As Byte, I As Long, II As Long, X As Long, Y As Long, Sprite As Long, spritetop As Long
    Dim tIcon As Long
    Dim rec As RECT
    Dim AttackSpeed As Long
    Dim npcNum As Long, QuestNum As Long
    Static tmrCount As Long
    Static setIt As Boolean
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    npcNum = MapNPC(MapNPCNum).num

    If npcNum = 0 Then Exit Sub ' No npc set
    
    Sprite = NPC(npcNum).Sprite

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub

    AttackSpeed = 1000

    ' Reset frame
    If MapNPC(MapNPCNum).Moving > 0 Then
        Anim = MapNPC(MapNPCNum).Step
    Else
        Anim = 0
    End If
    
    If NPC(npcNum).Animated = 0 Or MapNPC(MapNPCNum).Moving > 0 Then
        ' Check for attacking animation
        If MapNPC(MapNPCNum).AttackTimer + (AttackSpeed / 2) > timeGetTime Then
            If MapNPC(MapNPCNum).Attacking = 1 Then
                Anim = 3
            End If
        End If
    Else
        With MapNPC(MapNPCNum)
            If .AnimTimer + 100 <= timeGetTime Then
                .Anim = .Anim + 1
                If .Anim >= 4 Then .Anim = 0
                .AnimTimer = timeGetTime
            End If
            Anim = .Anim
        End With
    End If

    ' Check to see if we want to stop making him attack
    With MapNPC(MapNPCNum)
        If .AttackTimer + AttackSpeed < timeGetTime Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNPC(MapNPCNum).Dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
        Case DIR_UPLEFT
            spritetop = 1
        Case DIR_UPRIGHT
            spritetop = 2
        Case DIR_DOWNLEFT
            spritetop = 1
        Case DIR_DOWNRIGHT
            spritetop = 2
    End Select

    With rec
        .Top = (Tex_Character(Sprite).Height / 4) * spritetop
        .Bottom = .Top + Tex_Character(Sprite).Height / 4
        .Left = Anim * (Tex_Character(Sprite).Width / 4)
        .Right = .Left + (Tex_Character(Sprite).Width / 4)
    End With

    ' Calculate the X
    X = MapNPC(MapNPCNum).X * PIC_X + MapNPC(MapNPCNum).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNPC(MapNPCNum).Y * PIC_Y + MapNPC(MapNPCNum).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNPC(MapNPCNum).Y * PIC_Y + MapNPC(MapNPCNum).yOffset
    End If

    Call DrawSprite(Sprite, X, Y, rec)
    
    '///////////////////////////////////
'//////////QUEST DISPLAY////////////
'///////////////////////////////////

    'run through quests and see if this NPC can start/manage one.
    For I = 1 To MAX_QUESTS
        For II = 1 To Quest(I).Max_CLI
            If Quest(I).CLI(II).ItemIndex = npcNum Then 'this quest is assigned to this NPC
                'If its completed bypass rendering
                If Not Player(MyIndex).QuestCompleted(I) = 1 Then
                'show the icon based off quest status
                If Player(MyIndex).QuestCLI(I) = 0 Then 'haven't started this quest yet
                    If II = 1 Then 'make sure this is the first Greeter of the quest
                        tIcon = Quest_Icon_Start
                        GoTo PastQuestDataRetrieval
                    End If
                ElseIf Player(MyIndex).QuestCLI(I) = II Then 'make sure the player is on this NPC within the progress of the quest
                    tIcon = Quest_Icon_Progress
                    GoTo PastQuestDataRetrieval
                End If
                End If
            End If
        Next II
    Next I
    
PastQuestDataRetrieval:
    'display it above the NPC's name
    If NPC(npcNum).ShowQuestCompleteIcon = 1 Then tIcon = Quest_Icon_Finished
    
    QuestNum = DoesNPCStartQuest(npcNum)
    If QuestNum > 0 Then
        If Quest(QuestNum).CanBeRetaken = vbUnchecked Then
            If Player(MyIndex).QuestCompleted(QuestNum) = True Then
                tIcon = Quest_Icon_Completed
            End If
        End If
    End If
    
    If tIcon > 0 Then
        If timeGetTime >= tmrCount Then
            Call DrawQuestIcon(CLng(tIcon), MapNPCNum, True)
            If timeGetTime >= tmrCount + 500 Then setIt = False 'swap back
        Else
            Call DrawQuestIcon(CLng(tIcon), MapNPCNum) 'draw normal image.
        End If
        
        If Not setIt Then
            tmrCount = timeGetTime + 500
            setIt = True
        End If
    End If

'///////////////////////////////////
'//////////QUEST DISPLAY////////////
'///////////////////////////////////
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawNPC", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawPaperdoll(ByVal X2 As Long, ByVal Y2 As Long, ByVal Sprite As Long, ByVal Anim As Long, ByVal spritetop As Long)
    Dim rec As RECT
    Dim X As Long, Y As Long
    Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Sprite < 1 Or Sprite > NumPaperdolls Then Exit Sub
    
    With rec
        .Top = spritetop * (Tex_Paperdoll(Sprite).Height / 4)
        .Bottom = .Top + (Tex_Paperdoll(Sprite).Height / 4)
        .Left = Anim * (Tex_Paperdoll(Sprite).Width / 4)
        .Right = .Left + (Tex_Paperdoll(Sprite).Width / 4)
    End With
    
    ' Clipping
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)

    ' Clip to screen
    If Y < 0 Then
        With rec
            .Top = .Top - Y
        End With
        Y = 0
    End If

    If X < 0 Then
        With rec
            .Left = .Left - X
        End With
        X = 0
    End If
    
    RenderTexture Tex_Paperdoll(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawPaperdoll", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub DrawSprite(ByVal Sprite As Long, ByVal X2 As Long, Y2 As Long, rec As RECT)
    Dim X As Long
    Dim Y As Long
    Dim Width As Long
    Dim Height As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Sprite < 1 Or Sprite > NumCharacters Then Exit Sub
    
    X = ConvertMapX(X2)
    Y = ConvertMapY(Y2)
    Width = (rec.Right - rec.Left)
    Height = (rec.Bottom - rec.Top)
    
    RenderTexture Tex_Character(Sprite), X, Y, rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawAnimatedItems()
    Dim I As Long
    Dim ItemNum As Long, ItemPic As Long, Color As Long
    Dim X As Long, Y As Long
    Dim MaxFrames As Byte
    Dim amount As Long
    Dim rec As RECT, rec_pos As RECT
    Dim TmpItem As Long, AmountModifier As Long
    Dim NoRender(1 To MAX_INV) As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not InGame Then Exit Sub
    
    ' Check for map animation changes
    For I = 1 To MAX_MAP_ITEMS
        If MapItem(I).num > 0 Then
            ItemPic = Item(MapItem(I).num).Pic

            If ItemPic < 1 Or ItemPic > NumItems Then Exit Sub
            MaxFrames = Tex_Item(ItemPic).Width / PIC_X ' Work out how many frames there are.

            If MapItem(I).Frame < MaxFrames - 1 Then
                MapItem(I).Frame = MapItem(I).Frame + 1
            Else
                MapItem(I).Frame = 1
            End If
        End If
    Next
    
    For I = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, I)
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            AmountModifier = 0
            NoRender(I) = 0
            
            ' Exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    TmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                    If TradeYourOffer(X).num = I Then
                        ' Check if currency
                        If Not Item(TmpItem).stackable = 1 Then
                            ' Normal item don't render
                            NoRender(I) = 1
                        Else
                            ' If amount = all currency, remove from inventory
                            If TradeYourOffer(X).Value = GetPlayerInvItemValue(MyIndex, I) Then
                                NoRender(I) = 1
                            Else
                                ' Not all, change modifier to show change in currency count
                                AmountModifier = TradeYourOffer(X).Value
                            End If
                        End If
                    End If
                Next
            End If
                
            If NoRender(I) = 0 Then
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > PIC_X Then
                        MaxFrames = Tex_Item(ItemPic).Width / PIC_X ' Work out how many frames there are.
    
                        If InvItemFrame(I) < MaxFrames - 1 Then
                            InvItemFrame(I) = InvItemFrame(I) + 1
                        Else
                            InvItemFrame(I) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = Tex_Item(ItemPic).Width + (InvItemFrame(I) * 32) ' Middle to get the start of inv gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = InvLeft + ((InvOffsetX + PIC_X) * (((I - 1) Mod InvColumns)))
                            .Right = .Left + PIC_X
                        End With

                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
    
                        ' If item is a stack - draw the amount you have
                        If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                            Y = rec_pos.Top + 22
                            X = rec_pos.Left - 4
                            amount = GetPlayerInvItemValue(MyIndex, I) - AmountModifier
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If amount < 1000000 Then
                                Color = White
                            ElseIf amount > 1000000 And amount < 10000000 Then
                                Color = Yellow
                            ElseIf amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(amount), X, Y, Color
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If InBank Then
        For I = 1 To MAX_BANK
            ItemNum = GetBankItemNum(I)
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
    
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > PIC_X Then
                        MaxFrames = Tex_Item(ItemPic).Width / PIC_X ' Work out how many frames there are.
    
                        If BankItemFrame(I) < MaxFrames - 1 Then
                            BankItemFrame(I) = BankItemFrame(I) + 1
                        Else
                            BankItemFrame(I) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = Tex_Item(ItemPic).Width + (BankItemFrame(I) * 32) ' Middle to get the start of Bank gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = BankLeft + ((BankOffsetX + PIC_X) * (((I - 1) Mod BankColumns)))
                            .Right = .Left + PIC_X
                        End With
    
                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
    
                        ' If item is a stack - draw the amount you have
                        If GetBankItemValue(I) > 1 Then
                            Y = rec_pos.Top + 22
                            X = rec_pos.Left - 4
                            amount = GetBankItemValue(I)
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If amount < 1000000 Then
                                Color = White
                            ElseIf amount > 1000000 And amount < 10000000 Then
                                Color = Yellow
                            ElseIf amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(amount), X, Y, Color
                        End If
                    End If
                End If
            End If
        Next
    End If
    
    If InShop > 0 Then
        For I = 1 To MAX_TRADES
            ItemNum = Shop(InShop).TradeItem(I).Item
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
    
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > PIC_X Then
                        MaxFrames = Tex_Item(ItemPic).Width / PIC_X ' Work out how many frames there are.
    
                        If ShopItemFrame(I) < MaxFrames - 1 Then
                            ShopItemFrame(I) = ShopItemFrame(I) + 1
                        Else
                            ShopItemFrame(I) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = Tex_Item(ItemPic).Width + (ShopItemFrame(I) * 32) ' Middle to get the start of shop gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = ShopLeft + ((ShopOffsetX + PIC_X) * (((I - 1) Mod ShopColumns)))
                            .Right = .Left + PIC_X
                        End With
                        
                        ' If item is a stack - draw the amount you have
                        If Shop(InShop).TradeItem(I).ItemValue > 1 Then
                            Y = rec_pos.Top + 22
                            X = rec_pos.Left - 4
                            amount = Shop(InShop).TradeItem(I).ItemValue
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If amount < 1000000 Then
                                Color = White
                            ElseIf amount > 1000000 And amount < 10000000 Then
                                Color = Yellow
                            ElseIf amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(amount), X, Y, Color
                        End If
    
                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                    End If
                End If
            End If
        Next
    End If
    
    If frmMain.picTrade.Visible = True Then
        For I = 1 To MAX_INV
            ItemNum = TradeTheirOffer(I).num
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
    
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > PIC_X Then
                        MaxFrames = Tex_Item(ItemPic).Width / PIC_X ' Work out how many frames there are.
    
                        If InvItemFrame(I) < MaxFrames - 1 Then
                            InvItemFrame(I) = InvItemFrame(I) + 1
                        Else
                            InvItemFrame(I) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = Tex_Item(ItemPic).Width + (InvItemFrame(I) * 32) ' Middle to get the start of inv gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = InvTop - 12 + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = InvLeft + ((InvOffsetX + PIC_X) * (((I - 1) Mod InvColumns)))
                            .Right = .Left + PIC_X
                        End With

                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                    
                        ' If item is a stack - draw the amount you have
                        If TradeTheirOffer(I).Value > 1 Then
                            Y = rec_pos.Top + 22
                            X = rec_pos.Left - 4
                            amount = TradeTheirOffer(I).Value
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If amount < 1000000 Then
                                Color = White
                            ElseIf amount > 1000000 And amount < 10000000 Then
                                Color = Yellow
                            ElseIf amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(amount), X, Y, Color
                        End If
                    End If
                End If
            End If
        Next
        
         For I = 1 To MAX_INV
            ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)
            
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                ItemPic = Item(ItemNum).Pic
    
                If ItemPic > 0 And ItemPic <= NumItems Then
                    If Tex_Item(ItemPic).Width > PIC_X Then
                        MaxFrames = Tex_Item(ItemPic).Width / PIC_X ' Work out how many frames there are.
    
                        If InvItemFrame(I) < MaxFrames - 1 Then
                            InvItemFrame(I) = InvItemFrame(I) + 1
                        Else
                            InvItemFrame(I) = 1
                        End If
    
                        With rec
                            .Top = 0
                            .Bottom = 32
                            .Left = Tex_Item(ItemPic).Width + (InvItemFrame(I) * 32) ' Middle to get the start of inv gfx, then +32 for each frame
                            .Right = .Left + 32
                        End With
    
                        With rec_pos
                            .Top = InvTop - 12 + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = InvLeft + ((InvOffsetX + PIC_X) * (((I - 1) Mod InvColumns)))
                            .Right = .Left + PIC_X
                        End With

                        ' We'll now re-Draw the item, and place the currency value over it again :P
                        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                        
                        ' If item is a stack - draw the amount you have
                        If TradeYourOffer(I).Value > 1 Then
                            Y = rec_pos.Top + 22
                            X = rec_pos.Left - 4
                            amount = TradeYourOffer(I).Value
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If amount < 1000000 Then
                                Color = White
                            ElseIf amount > 1000000 And amount < 10000000 Then
                                Color = Yellow
                            ElseIf amount > 10000000 Then
                                Color = BrightGreen
                            End If
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            RenderText Font_Default, ConvertCurrency(amount), X, Y, Color
                        End If
                    End If
                End If
            End If
        Next
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawAnimatedItems", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawPlayerCharFace()
    Dim rec As RECT, rec_pos As RECT, FaceNum As Long, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If NumFaces = 0 Then Exit Sub
    
    FaceNum = Player(MyIndex).Face
    
    If FaceNum <= 0 Or FaceNum > NumFaces Then Exit Sub

    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = 100
        .Left = 0
        .Right = 100
    End With

    RenderTextureByRects Tex_Face(FaceNum), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmMain.picFace.Width
        .Y1 = 0
        .Y2 = frmMain.picFace.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmMain.picFace.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawPlayerCharFace", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub RefreshUpperTilesCacheWhole()
    Dim bbf As Direct3DSurface8
    Dim bUpper As Direct3DSurface8
    Dim X As Long, Y As Long

    Set upperTilesCache = Direct3DX.CreateTexture(Direct3D_Device, PIC_X * Map.MaxX + 32, PIC_Y * Map.MaxY + 32, D3DX_DEFAULT, D3DUSAGE_RENDERTARGET, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT)
    Set bbf = Direct3D_Device.GetRenderTarget
    Set bUpper = upperTilesCache.GetSurfaceLevel(0)

    Direct3D_Device.SetRenderTarget bUpper, Nothing, 0
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
    On Error Resume Next
    Direct3D_Device.BeginScene

    For X = 0 To (Map.MaxX)
        For Y = 0 To (Map.MaxY)
    
            If IsValidMapPoint(X, Y) Then
                Call DrawWholeMapUpperTiles(X, Y)
            End If
    
        Next
    Next

    Call Direct3D_Device.EndScene
    Direct3D_Device.SetRenderTarget bbf, Nothing, 0

    Set bbf = Nothing
    Set bUpper = Nothing
    
    Exit Sub
    ' Error handler
ErrorHandler:
    HandleError "RefreshUpperTilesCacheWhole", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub RefreshLowerTilesCacheWhole()
    Dim bbf As Direct3DSurface8
    Dim bLower As Direct3DSurface8
    Dim X  As Long, Y As Long
    
    Set lowerTilesCache = Direct3DX.CreateTexture(Direct3D_Device, PIC_X * Map.MaxX + 32, PIC_Y * Map.MaxY + 32, D3DX_DEFAULT, D3DUSAGE_RENDERTARGET, D3DFMT_UNKNOWN, D3DPOOL_DEFAULT)
    Set bbf = Direct3D_Device.GetRenderTarget
    Set bLower = lowerTilesCache.GetSurfaceLevel(0)
    
    Direct3D_Device.SetRenderTarget bLower, Nothing, 0
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
    On Error Resume Next
    Direct3D_Device.BeginScene
     
    For X = 0 To (Map.MaxX)
        For Y = 0 To (Map.MaxY)

            If IsValidMapPoint(X, Y) Then
                Call DrawWholeMapLowerTiles(X, Y)
            End If

        Next
    Next
    Call Direct3D_Device.EndScene
    Direct3D_Device.SetRenderTarget bbf, Nothing, 0
    Set bbf = Nothing
    Set bLower = Nothing
    Exit Sub
    ' Error handler
ErrorHandler:
    HandleError "RefreshLowerTilesCacheWhole", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawQuestIcon(Image As Long, npcNum As Long, Optional ByVal Up As Boolean = False)
    Dim I As Long, X As Long, Y As Long
    Dim Height As Long, Width As Long
    Dim tmpDX As DX8TextureRec
    Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Sprite = NPC(MapNPC(npcNum).num).Sprite
    
    
    Select Case Image
        Case Quest_Icon_Start
            If Up Then tmpDX = Quest_Start_U Else tmpDX = Quest_Start
        Case Quest_Icon_Finished
            If Up Then tmpDX = Quest_Finished_U Else tmpDX = Quest_Finished
        Case Quest_Icon_Progress
            If Up Then tmpDX = Quest_Progress_U Else tmpDX = Quest_Progress
        Case Quest_Icon_Completed
            If Up Then tmpDX = Quest_Completed_U Else tmpDX = Quest_Completed
    End Select
    
    ' Calculate the X
    X = MapNPC(npcNum).X * PIC_X + MapNPC(npcNum).xOffset - ((Tex_Character(Sprite).Width / 4 - 32) / 2)
    
    ' Is the player's height more than 32..?
    If (Tex_Character(Sprite).Height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        Y = MapNPC(npcNum).Y * PIC_Y + MapNPC(npcNum).yOffset - ((Tex_Character(Sprite).Height / 4) - 32)
    Else
        ' Proceed as normal
        Y = MapNPC(npcNum).Y * PIC_Y + MapNPC(npcNum).yOffset
    End If
    
    X = ConvertMapX(X)
    Y = ConvertMapY(Y) - 40
    
    Width = 33
    Height = 33
    
    RenderTexture tmpDX, X, Y, 0, 0, Width, Height, Width, Height
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawQuestIcon", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawInventory()
    Dim I As Long, X As Long, Y As Long, ItemNum As Long, ItemPic As Long
    Dim amount As Long
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Color As Long
    Dim TmpItem As Long, AmountModifier As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    With rec
        .Top = 0
        .Bottom = Tex_Base.Height
        .Left = 0
        .Right = Tex_Base.Width
    End With
    
    With rec_pos
        .Top = 0
        .Bottom = frmMain.picInventory.Height
        .Left = 0
        .Right = frmMain.picInventory.Width
    End With

    RenderTextureByRects Tex_Base, rec, rec_pos

    For I = 1 To MAX_INV
        ItemNum = GetPlayerInvItemNum(MyIndex, I)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            AmountModifier = 0
            
            ' Exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For X = 1 To MAX_INV
                    TmpItem = GetPlayerInvItemNum(MyIndex, TradeYourOffer(X).num)
                    If TradeYourOffer(X).num = I Then
                        ' Check if currency
                        If Not Item(TmpItem).stackable = 1 Then
                            ' Normal item, exit out
                            GoTo NextLoop
                        Else
                            ' If amount = all currency, remove from inventory
                            If TradeYourOffer(X).Value = GetPlayerInvItemValue(MyIndex, I) Then
                                GoTo NextLoop
                            Else
                                ' Not all, change modifier to show change in currency count
                                AmountModifier = TradeYourOffer(X).Value
                            End If
                        End If
                    End If
                Next
            End If

            If ItemPic > 0 And ItemPic <= NumItems Then
                If Tex_Item(ItemPic).Width <= 32 Then ' More than 1 frame is handled by anim sub
                     With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With

                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

                    ' If item is a stack - draw the amount you have
                    If GetPlayerInvItemValue(MyIndex, I) > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        amount = GetPlayerInvItemValue(MyIndex, I) - AmountModifier
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If amount < 1000000 Then
                            Color = White
                        ElseIf amount > 1000000 And amount < 10000000 Then
                            Color = Yellow
                        ElseIf amount > 10000000 Then
                            Color = BrightGreen
                        End If
                        
                        RenderText Font_Default, Format$(ConvertCurrency(str$(amount)), "#,###,###,###"), X, Y, Color
                    End If
                End If
            End If
        End If
        
NextLoop:
    Next
    
    With srcRect
        .X1 = 0
        .X2 = frmMain.picInventory.Width
        .Y1 = 0
        .Y2 = frmMain.picInventory.Height
    End With
    
    With destRECT
        .X1 = 0
        .X2 = frmMain.picInventory.Width
        .Y1 = 0
        .Y2 = frmMain.picInventory.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picInventory.hWnd, ByVal (0)
    
    ' Update animated items
    DrawAnimatedItems
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawInventory", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawTrade()
    Dim I As Long, X As Long, Y As Long, ItemNum As Long, ItemPic As Long
    Dim amount As Long
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Color As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not InGame Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    For I = 1 To MAX_INV
        ' Draw your own offer
        ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(I).num)

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic

            If ItemPic > 0 And ItemPic <= NumItems Then
                If Tex_Item(ItemPic).Width <= 32 Then
                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With

                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With
    
                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
    
                    ' If item is a stack - draw the amount you have
                    If TradeYourOffer(I).Value > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        amount = TradeYourOffer(I).Value
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If amount < 1000000 Then
                            Color = White
                        ElseIf amount > 1000000 And amount < 10000000 Then
                            Color = Yellow
                        ElseIf amount > 10000000 Then
                            Color = BrightGreen
                        End If
                        
                        RenderText Font_Default, Format$(ConvertCurrency(str$(amount)), "#,###,###,###"), X, Y, Color
                    End If
                End If
            End If
        End If
    Next
    
    With srcRect
        .X1 = 0
        .X2 = .X1 + 193
        .Y1 = 0
        .Y2 = .Y1 + 246
    End With
                    
    With destRECT
        .X1 = 0
        .X2 = .X1 + 193
        .Y1 = 0
        .Y2 = 246 + .Y1
    End With
        
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picYourTrade.hWnd, ByVal (0)
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
         
    For I = 1 To MAX_INV
        ' Draw their offer
        ItemNum = TradeTheirOffer(I).num

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic

            If ItemPic > 0 And ItemPic <= NumItems Then
                If Tex_Item(ItemPic).Width <= 32 Then
                    With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With
    
                    With rec_pos
                        .Top = InvTop + ((InvOffsetY + 32) * ((I - 1) \ InvColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = InvLeft + ((InvOffsetX + 32) * (((I - 1) Mod InvColumns)))
                        .Right = .Left + PIC_X
                    End With
    
                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
    
                    ' If item is a stack - draw the amount you have
                    If TradeTheirOffer(I).Value > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        amount = TradeTheirOffer(I).Value
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If amount < 1000000 Then
                            Color = White
                        ElseIf amount > 1000000 And amount < 10000000 Then
                            Color = Yellow
                        ElseIf amount > 10000000 Then
                            Color = BrightGreen
                        End If
                        
                        RenderText Font_Default, Format$(ConvertCurrency(str$(amount)), "#,###,###,###"), X, Y, Color
                    End If
                End If
            End If
        End If
    Next
    
    With srcRect
        .X1 = 0
        .X2 = .X1 + 193
        .Y1 = 0
        .Y2 = .Y1 + 246
    End With
                    
    With destRECT
        .X1 = 0
        .X2 = .X1 + 193
        .Y1 = 0
        .Y2 = 246 + .Y1
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picTheirTrade.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawTrade", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawPlayerSpells()
    Dim I As Long, X As Long, Y As Long, SpellNum As Long, SpellIcon As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim amount As String
    Dim rec As RECT, rec_pos As RECT
    Dim Color As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    With rec
        .Top = 0
        .Bottom = Tex_Base.Height
        .Left = 0
        .Right = Tex_Base.Width
    End With
    
    With rec_pos
        .Top = 0
        .Bottom = frmMain.picSpells.Height
        .Left = 0
        .Right = frmMain.picSpells.Width
    End With

    RenderTextureByRects Tex_Base, rec, rec_pos

    For I = 1 To MAX_PLAYER_SPELLS
        SpellNum = PlayerSpells(I)

        If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
            SpellIcon = Spell(SpellNum).Icon

            If SpellIcon > 0 And SpellIcon <= NumSpellIcons Then
                With rec
                    .Top = 0
                    .Bottom = 32
                    .Left = 0
                    .Right = 32
                End With
                
                If Not SpellCD(I) = 0 Then
                    rec.Left = 32
                    rec.Right = 64
                End If

                With rec_pos
                    .Top = SpellTop + ((SpellOffsetY + 32) * ((I - 1) \ SpellColumns))
                    .Bottom = .Top + PIC_Y
                    .Left = SpellLeft + ((SpellOffsetX + 32) * (((I - 1) Mod SpellColumns)))
                    .Right = .Left + PIC_X
                End With

                RenderTextureByRects Tex_SpellIcon(SpellIcon), rec, rec_pos
            End If
        End If
    Next
    
    With srcRect
        .X1 = 0
        .X2 = frmMain.picSpells.Width
        .Y1 = 0
        .Y2 = frmMain.picSpells.Height
    End With
    
    With destRECT
        .X1 = 0
        .X2 = frmMain.picSpells.Width
        .Y1 = 0
        .Y2 = frmMain.picSpells.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picSpells.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawPlayerSpells", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawShop()
    Dim I As Long, X As Long, Y As Long, ItemNum As Long, ItemPic As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim amount As String
    Dim rec As RECT, rec_pos As RECT
    Dim Color As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not InGame Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    For I = 1 To MAX_TRADES
        ItemNum = Shop(InShop).TradeItem(I).Item
        
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic

            If ItemPic > 0 And ItemPic <= NumItems Then
                If Tex_Item(ItemPic).Width <= 32 Then
                     With rec
                        .Top = 0
                        .Bottom = 32
                        .Left = 0
                        .Right = 32
                    End With

                    With rec_pos
                        .Top = ShopTop + ((ShopOffsetY + 32) * ((I - 1) \ ShopColumns))
                        .Bottom = .Top + PIC_Y
                        .Left = ShopLeft + ((ShopOffsetX + 32) * (((I - 1) Mod ShopColumns)))
                        .Right = .Left + PIC_X
                    End With

                    RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos
                    
                    ' If item is a stack - draw the amount you have
                    If Shop(InShop).TradeItem(I).ItemValue > 1 Then
                        Y = rec_pos.Top + 22
                        X = rec_pos.Left - 4
                        amount = Shop(InShop).TradeItem(I).ItemValue
                        
                        ' Draw currency but with k, m, b etc. using a convertion function
                        If amount < 1000000 Then
                            Color = White
                        ElseIf amount > 1000000 And amount < 10000000 Then
                            Color = Yellow
                        ElseIf amount > 10000000 Then
                            Color = BrightGreen
                        End If
                        
                        RenderText Font_Default, ConvertCurrency(amount), X, Y, Color
                    End If
                End If
            End If
        End If
    Next
    
    With srcRect
        .X1 = ShopLeft
        .X2 = .X1 + 192
        .Y1 = ShopTop
        .Y2 = .Y1 + 211
    End With
                
    With destRECT
        .X1 = ShopLeft
        .X2 = .X1 + 192
        .Y1 = ShopTop
        .Y2 = 211 + .Y1
    End With
                
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picShopItems.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawShop", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawDraggedItem(ByVal X As Long, ByVal Y As Long, Optional ByVal IsHotbarSlot As Boolean = False)
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim ItemNum As Long, ItemPic As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsHotbarSlot Then
        ItemNum = Hotbar(DragHotbarSlot).Slot
    Else
        ItemNum = GetPlayerInvItemNum(MyIndex, DragInvSlot)
    End If

    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
    
        ItemPic = Item(ItemNum).Pic
        
        If ItemPic < 1 Or ItemPic > NumItems Then Exit Sub
        
        With rec
            .Top = 0
            .Bottom = 32
            .Left = 0
            .Right = 32
        End With

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

        With frmMain.picTempInv
            .Top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
        
        With srcRect
            .X1 = 0
            .X2 = 32
            .Y1 = 0
            .Y2 = 32
        End With
        
        With destRECT
            .X1 = 2
            .Y1 = 2
            .Y2 = .Y1 + 32
            .X2 = .X1 + 32
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmMain.picTempInv.hWnd, ByVal (0)
    End If
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "DrawDraggedItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawDraggedSpell(ByVal X As Long, ByVal Y As Long, Optional ByVal IsHotbarSlot As Boolean = False)
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim SpellNum As Long, SpellPic As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If IsHotbarSlot Then
        SpellNum = Hotbar(DragHotbarSlot).Slot
    Else
        SpellNum = PlayerSpells(DragSpellSlot)
    End If
    
     If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        SpellPic = Spell(SpellNum).Icon
        
        If SpellPic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .Top = 0
            .Bottom = 32
            .Left = 0
            .Right = 32
        End With
        
        If IsHotbarSlot = False Then
            If SpellCD(DragSpellSlot) > 0 Then
                With rec
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = 32
                    .Right = .Left + PIC_X
                End With
            Else
                With rec
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = .Left + PIC_X
                End With
            End If
        Else
            If SpellCD(DragHotbarSpell) > 0 Then
                With rec
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = 32
                    .Right = .Left + PIC_X
                End With
            Else
                With rec
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = .Left + PIC_X
                End With
            End If
        End If

        With rec_pos
            .Top = 2
            .Bottom = .Top + PIC_Y
            .Left = 2
            .Right = .Left + PIC_X
        End With

        RenderTextureByRects Tex_SpellIcon(SpellPic), rec, rec_pos

        With frmMain.picTempSpell
            .Top = Y
            .Left = X
            .Visible = True
            .ZOrder (0)
        End With
        
        With srcRect
            .X1 = 0
            .X2 = 32
            .Y1 = 0
            .Y2 = 32
        End With
        
        With destRECT
            .X1 = 2
            .Y1 = 2
            .Y2 = .Y1 + 32
            .X2 = .X1 + 32
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmMain.picTempSpell.hWnd, ByVal (0)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawDraggedSpell", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawItemDesc(ByVal ItemNum As Long)
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim ItemPic As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
        ItemPic = Item(ItemNum).Pic

        If ItemPic = 0 Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        With rec
            .Top = 0
            .Bottom = 32
            .Left = 0
            .Right = 32
        End With

        With rec_pos
            .Top = 0
            .Bottom = 64
            .Left = 0
            .Right = 64
        End With
        
        RenderTextureByRects Tex_Item(ItemPic), rec, rec_pos

        With destRECT
            .X1 = 0
            .Y1 = 0
            .Y2 = 64
            .X2 = 64
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRECT, destRECT, frmMain.picItemDescPic.hWnd, ByVal (0)
    End If
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "DrawItemDesc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawSpellDesc(ByVal SpellNum As Long)
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim SpellPic As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        SpellPic = Spell(SpellNum).Icon

        If SpellPic <= 0 Or SpellPic > NumSpellIcons Then Exit Sub
        If LastSpellSlotDesc < 1 Or LastSpellSlotDesc > MAX_PLAYER_SPELLS Then Exit Sub
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene

        If SpellCD(LastSpellSlotDesc) > 0 Then
            With rec
                .Top = 0
                .Bottom = 32
                .Left = 32
                .Right = 64
            End With
        Else
            With rec
                .Top = 0
                .Bottom = 32
                .Left = 0
                .Right = 32
            End With
        End If
        
        With rec_pos
            .Top = 0
            .Bottom = .Top + 64
            .Left = 0
            .Right = .Left + 64
        End With
        
        RenderTextureByRects Tex_SpellIcon(SpellPic), rec, rec_pos

        With destRECT
            .X1 = 0
            .Y1 = 0
            .Y2 = 64
            .X2 = 64
        End With
        
        Direct3D_Device.EndScene
        Direct3D_Device.Present destRECT, destRECT, frmMain.picSpellDescPic.hWnd, ByVal (0)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawSpellDesc", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawTileOutline()
    Dim rec As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Map.OptBlock.Value Then Exit Sub

    With rec
        .Top = 0
        .Bottom = .Top + PIC_Y
        .Left = 0
        .Right = .Left + PIC_X
    End With

    RenderTexture Tex_Misc, ConvertMapX(CurX * PIC_X), ConvertMapY(CurY * PIC_Y), rec.Left, rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, rec.Right - rec.Left, rec.Bottom - rec.Top, D3DColorRGBA(255, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawTileOutline", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub Menu_DrawCharacter()
    Dim Sprite As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim sRECT As RECT
    Dim dRect As RECT
    Dim Width As Long, Height As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If frmMenu.cmbClass.ListIndex = -1 Then Exit Sub
    
    If frmMenu.optMale.Value = True Then
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).MaleSprite
    Else
        Sprite = Class(frmMenu.cmbClass.ListIndex + 1).FemaleSprite
    End If
    
    If Sprite < 1 Or Sprite > NumCharacters Then
        frmMenu.picSprite.Cls
        Exit Sub
    End If
    
    Width = Tex_Character(Sprite).Width / 4
    Height = Tex_Character(Sprite).Height / 4
    
    frmMenu.picSprite.Width = Width
    frmMenu.picSprite.Height = Height
    
    sRECT.Top = 0
    sRECT.Bottom = sRECT.Top + Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Width
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    RenderTextureByRects Tex_Character(Sprite), sRECT, dRect
    
    With srcRect
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    With destRECT
        .X1 = 0
        .X2 = Width
        .Y1 = 0
        .Y2 = Height
    End With
                    
    Direct3D_Device.EndScene
    On Error Resume Next
    Direct3D_Device.Present srcRect, destRECT, frmMenu.picSprite.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Menu_DrawCharacter", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub Render_Graphics()
    Dim X As Long
    Dim Y As Long
    Dim I As Long
    Dim rec As RECT
    Dim rec_pos As RECT, srcRect As D3DRECT, LocY As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for device lost
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then HandleDeviceLost: Exit Sub
    
    ' Don't render
    If frmMain.WindowState = vbMinimized Then Exit Sub
    
    For I = 1 To NumTextures

        If gTexture(I).Timer < timeGetTime And gTexture(I).Timer <> 0 Then
            UnsetTexture I
        End If

    Next
    
    If GettingMap Then Exit Sub
    
    ' Update the viewpoint
    Call UpdateCamera

    If redrawMapCache Then
        RefreshLowerTilesCacheWhole
        RefreshUpperTilesCacheWhole
        redrawMapCache = False
        
        If FormVisible("frmMapPreview") Then frmMapPreview.RecalcuateDimensions
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorARGB(0, 0, 0, 0), 1#, 0
    On Error Resume Next
    Direct3D_Device.BeginScene
 
    ' Update draw Name
    UpdateDrawMapName
    
    ' Draw panorama
    If NumPanoramas > 0 Then
        If Map.Panorama > 0 And Map.Panorama <= NumPanoramas Then
            RenderTexture Tex_Panorama(Map.Panorama), 0, 0, 0, 0, Tex_Panorama(Map.Panorama).Width, Tex_Panorama(Map.Panorama).Height, Tex_Panorama(Map.Panorama).Width, Tex_Panorama(Map.Panorama).Height, -1
        End If
    End If
    
    ' Draw lower tiles
    RenderCache lowerTilesCache, 0, 0, TileView.Left * PIC_X + Camera.Left, TileView.Top * PIC_Y + Camera.Top, ScreenX, ScreenY, ScreenX, ScreenY
     
    ' Render the decals
    If Options.Blood = 1 Then
        For I = 1 To Blood_HighIndex
            Call DrawBlood(I)
        Next
    End If

    ' Draw out the items
    If NumItems > 0 Then
        For I = 1 To MAX_MAP_ITEMS
            If MapItem(I).num > 0 Then
                Call DrawMapItem(I)
            End If
        Next
    End If
    
    ' Draw out lower events
    If Map.CurrentEvents > 0 Then
        For I = 1 To Map.CurrentEvents
            If Map.MapEvents(I).Position = 0 Then
                DrawEvent I
            End If
        Next
    End If
    
    ' Draw animations
    If NumAnimations > 0 Then
        For I = 1 To MAX_BYTE
            If AnimInstance(I).Used(0) Then
                DrawAnimation I, 0
            End If
        Next
    End If

    ' Y-based render. Renders players, npcs, and resources based on Y-axis.
    For Y = TileView.Top To TileView.Bottom
        ' NPCs
        For I = 1 To Map.NPC_HighIndex
            If MapNPC(I).Y = Y Then
                Call DrawNPC(I)
            End If
        Next
        
        ' Players
        For I = 1 To Player_HighIndex
            If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
                If Player(I).Y = Y Then
                    If Not I = MyIndex Then
                        Call DrawPlayer(I)
                    End If
                End If
            End If
        Next
        
        ' Render our sprite now so it's always at the top
        If Player(MyIndex).Y = Y Then
            Call DrawPlayer(MyIndex)
        End If
        
        ' Events
        If Map.CurrentEvents > 0 Then
            For I = 1 To Map.CurrentEvents
                If Map.MapEvents(I).Position = 1 Then
                    If Y = Map.MapEvents(I).Y Then
                        DrawEvent I
                    End If
                End If
            Next
        End If
        
        ' Resources
        If NumResources > 0 Then
            If Resources_Init Then
                If Resource_Index > 0 Then
                    For I = 1 To Resource_Index
                        If MapResource(I).Y = Y Then
                            Call DrawMapResource(I)
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    ' Animations
    If NumAnimations > 0 Then
        For I = 1 To MAX_BYTE
            If AnimInstance(I).Used(1) Then
                DrawAnimation I, 1
            End If
        Next
    End If
    
    'Draw out upper tiles
    RenderCache upperTilesCache, 0, 0, TileView.Left * PIC_X + Camera.Left, TileView.Top * PIC_Y + Camera.Top, ScreenX, ScreenY, ScreenX, ScreenY

    ' Tile preview
    If InMapEditor And Not displayTilesets Then
        If frmMain.chkTilePreview.Value And frmEditor_Map.chkRandom = 0 And frmEditor_Map.scrlAutotile.Value = 0 And frmEditor_Map.OptLayers.Value Then
            Call EditorMap_DrawTilePreview
        End If
    End If
    
    ' Draw out higher events
    If Map.CurrentEvents > 0 Then
        For I = 1 To Map.CurrentEvents
            If Map.MapEvents(I).Position = 2 Then
                DrawEvent I
            End If
        Next
    End If
    
    DrawWeather
    DrawFog
    DrawTint
    
    ' Draw the bars
    Call DrawBars
    
    ' Draw out a square at the mouse cursor
    If InMapEditor Then
        If frmEditor_Map.OptBlock.Value Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.Bottom
                    If IsValidMapPoint(X, Y) Then
                        Call DrawGrid(X, Y)
                        Call DrawDirection(X, Y)
                    End If
                Next
            Next
        ElseIf frmMain.chkGrid Then
            For X = TileView.Left To TileView.Right
                For Y = TileView.Top To TileView.Bottom
                    If IsValidMapPoint(X, Y) Then
                        Call DrawGrid(X, Y)
                    End If
                Next
            Next
        End If
    End If
    
    ' Draw the target icon
    If MyTarget > 0 Then
        If MyTargetType = TARGET_TYPE_PLAYER Then
            DrawTarget (Player(MyTarget).X * 32) + TempPlayer(MyTarget).xOffset, (Player(MyTarget).Y * 32) + TempPlayer(MyTarget).yOffset
        ElseIf MyTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNPC(MyTarget).X * 32) + MapNPC(MyTarget).xOffset, (MapNPC(MyTarget).Y * 32) + MapNPC(MyTarget).yOffset
        End If
    End If
    
    ' Draw the hover icon
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If Player(I).Map = Player(MyIndex).Map Then
                If CurX = Player(I).X And CurY = Player(I).Y Then
                    If MyTargetType = TARGET_TYPE_PLAYER And MyTarget = I Then
                        ' Don't render
                    Else
                        DrawHover TARGET_TYPE_PLAYER, I, (Player(I).X * 32) + TempPlayer(I).xOffset, (Player(I).Y * 32) + TempPlayer(I).yOffset
                    End If
                End If
            End If
        End If
    Next
    
    For I = 1 To Map.NPC_HighIndex
        If MapNPC(I).num > 0 Then
            If CurX = MapNPC(I).X And CurY = MapNPC(I).Y Then
                If MyTargetType = TARGET_TYPE_NPC And MyTarget = I Then
                    ' Don't render
                Else
                    DrawHover TARGET_TYPE_NPC, I, (MapNPC(I).X * 32) + MapNPC(I).xOffset, (MapNPC(I).Y * 32) + MapNPC(I).yOffset
                End If
            End If
        End If
    Next
    
    ' Draw weater
    If DrawThunder > 0 Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, 160): DrawThunder = DrawThunder - 1
    
    ' Get rec
    With rec
        .Top = Camera.Top
        .Bottom = .Top + ScreenY
        .Left = Camera.Left
        .Right = .Left + ScreenX
    End With
        
    ' rec_pos
    With rec_pos
        .Bottom = ScreenY
        .Right = ScreenX
    End With
        
    With srcRect
        .X1 = 0
        .X2 = frmMain.picScreen.ScaleWidth
        .Y1 = 0
        .Y2 = frmMain.picScreen.ScaleHeight
    End With
    
    If InMapEditor Then Call DrawMapAttributes
    
    ' Draw player names
    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerMap(I) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(I)
        End If
    Next
    
    For I = 1 To Map.CurrentEvents
        If Map.MapEvents(I).Visible = 1 Then
            If Map.MapEvents(I).ShowName = 1 Then
                Call DrawEventName(I)
            End If
        End If
    Next
    
    ' Draw npc names
    For I = 1 To Map.NPC_HighIndex
        If MapNPC(I).num > 0 Then
            Call DrawNPCName(I)
        End If
    Next
    
    ' draw the messages
    For I = 1 To ChatBubble_HighIndex
        If ChatBubble(I).active Then
            Call DrawChatBubble(I)
        End If
    Next
    
    ' Draw emotions
    DrawEmoticons
    
    ' Draw action messages
    For I = 1 To Action_HighIndex
        Call DrawActionMsg(I)
    Next
    
    ' Draw map name
    RenderText Font_Default, Map.Name, DrawMapNameX, DrawMapNameY, DrawMapNameColor
    
    If InMapEditor And (frmEditor_Map.OptEvents.Value Or frmMain.chkDrawEvents.Value) Then DrawEvents
    If InMapEditor And frmEditor_Map.OptLayers And Not displayTilesets Then DrawTileOutline

    If FadeAmount > 0 Then RenderTexture Tex_Fade, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, D3DColorRGBA(255, 255, 255, FadeAmount)
    If FlashTimer > timeGetTime Then RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, -1
    
    ' Draw fps
    If BFPS Then
        LocY = 12
        If GUIVisible Then
            RenderText Font_Default, "FPS: " & GameFPS & " Ping: " & CStr(Ping), 12, 104, White
        Else
            RenderText Font_Default, "FPS: " & GameFPS & " Ping: " & CStr(Ping), 12, 24, White
        End If
    End If
    
    ' Draw loc
    If BLoc Then
        If GUIVisible Then
            RenderText Font_Default, Trim$("Cur X: " & CurX & " Y: " & CurY), 12, 108 + LocY, Yellow
            RenderText Font_Default, Trim$("Loc X: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 12, 120 + LocY, Yellow
            RenderText Font_Default, Trim$(" (Map #" & GetPlayerMap(MyIndex) & ")"), 12, 132 + LocY, Yellow
        Else
            RenderText Font_Default, Trim$("Cur X: " & CurX & " Y: " & CurY), 12, 28 + LocY, Yellow
            RenderText Font_Default, Trim$("Loc X: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), 12, 40 + LocY, Yellow
            RenderText Font_Default, Trim$(" (Map #" & GetPlayerMap(MyIndex) & ")"), 12, 52 + LocY, Yellow
        End If
    End If
    
    If FormVisible("frmEditor_Map") And displayTilesets Then
        EditorMap_DrawTileset
        ' Tiles preview
        If frmEditor_Map.scrlAutotile.Value = 0 And frmEditor_Map.OptLayers.Value Then
            Call EditorMap_DrawTilePreview
        End If
    End If
    Call Direct3D_Device.EndScene
    On Error Resume Next
    Call Direct3D_Device.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    
    If Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST Or Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICENOTRESET Then
        HandleDeviceLost

        Exit Sub

    Else
        On Error Resume Next
    
        Direct3D_Device.Present ByVal 0, ByVal 0, 0, ByVal 0
        DrawGDI
    End If
    Exit Sub
    
 '  Error Handler
ErrorHandler:
    HandleError "Render_Graphics", "modGraphics", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function ConvertMapX(ByVal X As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ConvertMapX = X - (TileView.Left * PIC_X) - Camera.Left
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ConvertMapX", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ConvertMapY(ByVal Y As Long) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ConvertMapY = Y - (TileView.Top * PIC_Y) - Camera.Top
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ConvertMapY", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function InViewPort(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If X < TileView.Left Then Exit Function
    If Y < TileView.Top Then Exit Function
    If X > TileView.Right Then Exit Function
    If Y > TileView.Bottom Then Exit Function
    
    InViewPort = True
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "InViewPort", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function IsValidMapPoint(ByVal X As Long, ByVal Y As Long) As Boolean
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If X < 0 Then Exit Function
    If Y < 0 Then Exit Function
    If X > Map.MaxX Then Exit Function
    If Y > Map.MaxY Then Exit Function
    
    IsValidMapPoint = True
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsValidMapPoint", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub LoadTilesets()
    Dim X As Long
    Dim Y As Long
    Dim I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ReDim TilesetInUse(0 To NumTileSets)
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For I = 1 To MapLayer.Layer_Count - 1
                ' Check exists
                If Map.Tile(X, Y).Layer(I).Tileset > 0 And Map.Tile(X, Y).Layer(I).Tileset <= NumTileSets Then
                    TilesetInUse(Map.Tile(X, Y).Layer(I).Tileset) = True
                End If
            Next
        Next
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "LoadTilesets", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawBank()
    Dim I As Long, X As Long, Y As Long, ItemNum As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim amount As String
    Dim sRECT As RECT, dRect As RECT
    Dim Sprite As Long, Color As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmMain.picBank.Visible Then
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
    
        For I = 1 To MAX_BANK
            ItemNum = GetBankItemNum(I)
            If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
                Sprite = Item(ItemNum).Pic
                
                If Sprite > 0 Or Sprite <= NumItems Then
                    If Tex_Item(Sprite).Width <= 32 Then
                        With sRECT
                            .Top = 0
                            .Bottom = 32
                            .Left = 0
                            .Right = 32
                        End With
    
                        With dRect
                            .Top = BankTop + ((BankOffsetY + 32) * ((I - 1) \ BankColumns))
                            .Bottom = .Top + PIC_Y
                            .Left = BankLeft + ((BankOffsetX + 32) * (((I - 1) Mod BankColumns)))
                            .Right = .Left + PIC_X
                        End With
                        
                        RenderTextureByRects Tex_Item(Sprite), sRECT, dRect
    
                        ' If item is a stack - draw the amount you have
                        If GetBankItemValue(I) > 1 Then
                            Y = dRect.Top + 22
                            X = dRect.Left
                            amount = GetBankItemValue(I)
                            
                            ' Draw currency but with k, m, b etc. using a convertion function
                            If CLng(amount) < 1000000 Then
                                Color = White
                            ElseIf CLng(amount) > 1000000 And CLng(amount) < 10000000 Then
                                Color = Yellow
                            ElseIf CLng(amount) > 10000000 Then
                                Color = BrightGreen
                            End If
                            RenderText Font_Default, ConvertCurrency(amount), X, Y, Color
                        End If
                    End If
                End If
            End If
        Next
    
        With srcRect
            .X1 = BankLeft
            .X2 = .X1 + 400
            .Y1 = BankTop
            .Y2 = .Y1 + 310
        End With
                    
        With destRECT
            .X1 = BankLeft
            .X2 = .X1 + 400
            .Y1 = BankTop
            .Y2 = 310 + .Y1
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmMain.picBank.hWnd, ByVal (0)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawBank", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawBankItem(ByVal X As Long, ByVal Y As Long)
    Dim sRECT As RECT, dRect As RECT, srcRect As D3DRECT, destRECT As D3DRECT
    Dim ItemNum As Long
    Dim Sprite As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ItemNum = GetBankItemNum(DragBankSlot)
    Sprite = Item(GetBankItemNum(DragBankSlot)).Pic
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 255), 1#, 0
    Direct3D_Device.BeginScene
    
    If ItemNum > 0 Then
        If ItemNum <= MAX_ITEMS Then
            With sRECT
                 .Top = 0
                .Bottom = 32
                .Left = 0
                .Right = 32
            End With
        End If
    End If
    
    With dRect
        .Top = 2
        .Bottom = .Top + PIC_Y
        .Left = 2
        .Right = .Left + PIC_X
    End With

    RenderTextureByRects Tex_Item(Sprite), sRECT, dRect
    
    With frmMain.picTempBank
        .Top = Y
        .Left = X
        .Visible = True
        .ZOrder (0)
    End With
    
    With srcRect
        .X1 = 0
        .X2 = 32
        .Y1 = 0
        .Y2 = 32
    End With
    
    With destRECT
        .X1 = 2
        .Y1 = 2
        .Y2 = .Y1 + 32
        .X2 = .X1 + 32
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, destRECT, frmMain.picTempBank.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawBankItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal Alpha As Byte = 255)
    Dim yOffset As Long, xOffset As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Calculate the offset
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    RenderTexture Tex_Tileset(Map.Tile(X, Y).Layer(layerNum).Tileset), destX, destY, Autotile(X, Y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(X, Y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, D3DColorARGB(Alpha, 255, 255, 255)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawAutoTile", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorMap_DrawRandom()
    Dim sRECT As RECT
    Dim dRect As RECT
    Dim X As Long, Y As Long
    Dim I As Byte
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    For I = 0 To 3
        If RandomTileSheet(I) = 0 Then
            Exit Sub
        End If
        
        X = RandomTile(I) Mod 16
        Y = (RandomTile(I) - X) / 16
        
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        
        sRECT.Top = Y * PIC_Y
        sRECT.Bottom = sRECT.Top + PIC_Y
        sRECT.Left = X * PIC_X
        sRECT.Right = sRECT.Left + PIC_X
        
        dRect = sRECT
        dRect.Top = 0
        dRect.Bottom = PIC_Y
        dRect.Left = 0
        dRect.Right = PIC_X
        
        RenderTextureByRects Tex_Tileset(RandomTileSheet(I)), sRECT, dRect
    
        Direct3D_Device.EndScene
        Direct3D_Device.Present dRect, dRect, frmEditor_Map.picRandomTile(I).hWnd, ByVal (0)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorMap_DrawRandom", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' Character Editor
Public Sub EditorChar_AnimSprite(container As PictureBox, SpriteNum As String, ByRef spritePosition As Byte)
    Dim srcRect As D3DRECT, destRECT As D3DRECT
    Dim sRECT As RECT
    Dim dRect As RECT
    Dim X As Byte, Y As Byte
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If spritePosition > 15 Then spritePosition = 0
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    X = spritePosition Mod 4
    Y = (spritePosition - X) / 4
    
    sRECT.Top = Y * 48
    sRECT.Bottom = sRECT.Top + 48
    sRECT.Left = X * 32
    sRECT.Right = sRECT.Left + 32

    dRect = sRECT
    dRect.Top = 0
    dRect.Bottom = 48
    dRect.Left = 0
    dRect.Right = 32
    
    RenderTextureByRects Tex_Character(CLng(SpriteNum)), sRECT, dRect
          
    With destRECT
        .X1 = 0
        .X2 = container.ScaleWidth
        .Y1 = 0
        .Y2 = container.ScaleHeight
    End With

    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, container.hWnd, ByVal (0)

    spritePosition = spritePosition + 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorChar_AnimSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
' Character Editor
Public Sub drawRecentItem(SpriteNum As Integer)
    'Dim destRECT As D3DRECT
    Dim drawRect As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    

    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    drawRect.Top = 0
    drawRect.Bottom = 32
    drawRect.Left = 0
    drawRect.Right = 32

    RenderTextureByRects Tex_Item(SpriteNum), drawRect, drawRect

    Direct3D_Device.EndScene
    Direct3D_Device.Present drawRect, drawRect, frmAdmin.picRecentItem.hWnd, ByVal (0)

    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorChar_AnimSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorClass_DrawFace(ByVal Gender As Byte)
    Dim Sprite As Long
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If NumFaces = 0 Then Exit Sub
    
    If Gender = 0 Then
        Sprite = frmEditor_Class.scrlMFace.Value
    Else
        Sprite = frmEditor_Class.scrlFFace.Value
    End If
    
    If Sprite <= 0 Or Sprite > NumFaces Then
        frmEditor_Class.picFace.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Bottom = Tex_Face(Sprite).Height
        .Left = 0
        .Right = Tex_Face(Sprite).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = Tex_Face(Sprite).Height
        .Left = 0
        .Right = Tex_Face(Sprite).Width
    End With

    RenderTextureByRects Tex_Face(Sprite), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Class.picFace.Width
        .Y1 = 0
        .Y2 = frmEditor_Class.picFace.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmEditor_Class.picFace.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorClass_DrawFace", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawEventChatFace()
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If NumFaces = 0 Then Exit Sub
    
    If EventFace <= 0 Or EventFace > NumFaces Then Exit Sub
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = Tex_Face(EventFace).Height
        .Left = 0
        .Right = Tex_Face(EventFace).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = 96
        .Left = 0
        .Right = 96
    End With

    RenderTextureByRects Tex_Face(EventFace), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmMain.picChatFace.Width
        .Y1 = 0
        .Y2 = frmMain.picChatFace.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmMain.picChatFace.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawEventChatFace", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub EditorEvent_DrawFace()
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT
    Dim FaceNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If NumFaces = 0 Then Exit Sub
    
    FaceNum = frmEditor_Events.scrlFace.Value
    
    If FaceNum <= 0 Or FaceNum > NumFaces Then
        frmEditor_Events.picFace.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    RenderTextureByRects Tex_Face(FaceNum), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Events.picFace.Width
        .Y1 = 0
        .Y2 = frmEditor_Events.picFace.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmEditor_Events.picFace.hWnd, ByVal (0)
    
    frmEditor_Events.picFace.Height = PixelsToTwips(Tex_Face(FaceNum).Height, 1)
    frmEditor_Events.picFace.Width = PixelsToTwips(Tex_Face(FaceNum).Width, 0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorEvent_DrawFace", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub EditorEvent_DrawFace2()
    Dim rec As RECT, rec_pos As RECT, srcRect As D3DRECT
    Dim FaceNum As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If NumFaces = 0 Then Exit Sub
    
    FaceNum = frmEditor_Events.scrlFace2.Value
    
    If FaceNum <= 0 Or FaceNum > NumFaces Then
        frmEditor_Events.picFace2.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene

    With rec
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    With rec_pos
        .Top = 0
        .Bottom = Tex_Face(FaceNum).Height
        .Left = 0
        .Right = Tex_Face(FaceNum).Width
    End With

    RenderTextureByRects Tex_Face(FaceNum), rec, rec_pos
    
    With srcRect
        .X1 = 0
        .X2 = frmEditor_Events.picFace2.Width
        .Y1 = 0
        .Y2 = frmEditor_Events.picFace2.Height
    End With
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present srcRect, srcRect, frmEditor_Events.picFace2.hWnd, ByVal (0)
    
    frmEditor_Events.picFace2.Height = PixelsToTwips(Tex_Face(FaceNum).Height, 1)
    frmEditor_Events.picFace2.Width = PixelsToTwips(Tex_Face(FaceNum).Width, 0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorEvent_DrawFace2", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorMap_DrawTilePreview()
    Dim Height As Long
    Dim Width As Long
    Dim X As Long
    Dim Y As Long
    Dim Tileset As Long
    Dim srcRect As RECT
    Dim destRECT As D3DRECT
    Dim dRect As RECT
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not IsInBounds Then Exit Sub
    
    ' Find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    If Not displayTilesets Then
        X = CurX * PIC_X
        Y = CurY * PIC_Y
    Else
        X = Tex_Tileset(frmEditor_Map.scrlTileSet.Value).Width + 40
        Y = 32
    End If
    
    Height = EditorTileHeight
    Width = EditorTileWidth
    
    With srcRect
        .Left = 0
        .Top = 0
        .Right = srcRect.Left - Width
        .Bottom = srcRect.Top - Height
    End With
    
    With destRECT
        .X1 = (EditorTileX * PIC_X) - srcRect.Left
        .X2 = (EditorTileWidth * PIC_X) + .X1
        .Y1 = (EditorTileY * PIC_Y) - srcRect.Top
        .Y2 = (EditorTileHeight * PIC_Y) + .Y1
    End With
    With dRect
        .Bottom = 0
        .Left = (Tex_Tileset(Tileset).Width) + (((ScreenX - (Tex_Tileset(Tileset).Width)) - 170) / 2)
        .Right = 0
        .Top = 10
    End With
    
    If Not displayTilesets Then
        RenderTexture Tex_Tileset(Tileset), ConvertMapX(X), ConvertMapY(Y), destRECT.X1, destRECT.Y1, Width * PIC_X, Height * PIC_Y, Width * PIC_X, Height * PIC_Y, D3DColorARGB(200, 255, 255, 255)
    Else
        'RenderTexture Tex_Tileset(Tileset), X, Y, destRECT.X1, destRECT.Y1, Width * PIC_X, Height * PIC_Y, Width * PIC_X, Height * PIC_Y
        'RenderText Font_Default, "PREVIEW OF SELECTED TILES", dRect.Left, dRect.Top, White
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawMapTilesPreview", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorClass_DrawSprite(ByVal Gender As Byte)
    Dim Sprite As Integer
    Dim sRECT As RECT
    Dim dRect As RECT
    Dim destRECT As D3DRECT

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Gender = 0 Then
        Sprite = frmEditor_Class.scrlMSprite.Value
    Else
        Sprite = frmEditor_Class.scrlFSprite.Value
    End If

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_Class.picSprite.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = Tex_Character(Sprite).Height / 4
    sRECT.Left = (Tex_Character(Sprite).Width / 4) * 2 ' Facing down
    sRECT.Right = sRECT.Left + Tex_Character(Sprite).Width / 4
    dRect.Top = 0
    dRect.Bottom = Tex_Character(Sprite).Height / 4
    dRect.Left = 0
    dRect.Right = Tex_Character(Sprite).Width / 4
    
    frmEditor_Class.picSprite.Width = PixelsToTwips(Tex_Character(Sprite).Width / 4, 0)
    frmEditor_Class.picSprite.Height = PixelsToTwips(Tex_Character(Sprite).Height / 4, 1)

    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRECT, dRect
    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_Class.picSprite.Width
        .Y1 = 0
        .Y2 = frmEditor_Class.picSprite.Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Class.picSprite.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorClass_DrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub RenderOptionButton(ByRef ThePictureBox As PictureBox, ByVal TheOption As Byte, ByVal TheValue As Byte)
    Dim FileName As String

    If TheValue = 0 Then
        FileName = App.Path & GFX_PATH & "gui/main/buttons/option_off.bmp"
    ElseIf TheValue = 1 Then
        FileName = App.Path & GFX_PATH & "gui/main/buttons/option_on.bmp"
    ElseIf TheValue = 2 Then
        FileName = App.Path & GFX_PATH & "gui/main/buttons/option_off_hover.bmp"
    ElseIf TheValue = 3 Then
        FileName = App.Path & GFX_PATH & "gui/main/buttons/option_on_hover.bmp"
    End If
    
    OptionButton(TheOption).State = TheValue
    ThePictureBox.Picture = LoadPicture(FileName)
End Sub

Public Sub ResizeHPBar()
    Dim MoveSpeed As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If CurrentHPBarWidth = 0 Then Exit Sub

    If Not CurrentHPBarWidth = NewHPBarWidth Then
        If NewHPBarWidth > OldHPBarWidth Then
            MoveSpeed = (NewHPBarWidth - OldHPBarWidth) / 25

            If MoveSpeed < 1 Then MoveSpeed = 1

            If CurrentHPBarWidth + MoveSpeed > NewHPBarWidth Then
                CurrentHPBarWidth = NewHPBarWidth
            Else
                CurrentHPBarWidth = CurrentHPBarWidth + MoveSpeed
            End If

        Else
            MoveSpeed = (OldHPBarWidth - NewHPBarWidth) / 25

            If MoveSpeed < 1 Then MoveSpeed = 1

            If CurrentHPBarWidth + MoveSpeed < NewHPBarWidth Then
                CurrentHPBarWidth = NewHPBarWidth
            Else
                CurrentHPBarWidth = CurrentHPBarWidth - MoveSpeed
            End If
        
        End If
        
        If CurrentHPBarWidth >= 0 Then frmMain.imgHPBar.Width = CurrentHPBarWidth
    End If
    Exit Sub
    
' Error Handler
ErrorHandler:
    HandleError "ResizeHPBar", "modGeneral", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
    
End Sub

Public Sub ResizeMPBar()
    Dim MoveSpeed As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If CurrentMPBarWidth = 0 Then Exit Sub

    If Not CurrentMPBarWidth = NewMPBarWidth Then
        If NewMPBarWidth > OldMPBarWidth Then
            MoveSpeed = (NewMPBarWidth - OldMPBarWidth) / 25

            If MoveSpeed < 1 Then MoveSpeed = 1

            If CurrentMPBarWidth + MoveSpeed > NewMPBarWidth Then
                CurrentMPBarWidth = NewMPBarWidth
            Else
                CurrentMPBarWidth = CurrentMPBarWidth + MoveSpeed
            End If

        Else
            MoveSpeed = (OldMPBarWidth - NewMPBarWidth) / 25

            If MoveSpeed < 1 Then MoveSpeed = 1

            If CurrentMPBarWidth + MoveSpeed < NewMPBarWidth Then
                CurrentMPBarWidth = NewMPBarWidth
            Else
                CurrentMPBarWidth = CurrentMPBarWidth - MoveSpeed
            End If
            
        End If
        
        If CurrentMPBarWidth >= 0 Then frmMain.imgMPBar.Width = CurrentMPBarWidth
    End If
   Exit Sub
   
'' Error Handler
ErrorHandler:
    HandleError "ResizeMPBar", "modGeneral", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ResizeExpBar()
    Dim MoveSpeed As Long
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If CurrentMPBarWidth = 0 Then Exit Sub

    If Not CurrentEXPBarWidth = NewEXPBarWidth Then
        If NewEXPBarWidth > OldEXPBarWidth Then
            MoveSpeed = (NewEXPBarWidth - OldEXPBarWidth) / 25

            If MoveSpeed < 1 Then MoveSpeed = 1

            If CurrentEXPBarWidth + MoveSpeed > NewEXPBarWidth Then
                CurrentEXPBarWidth = NewEXPBarWidth
            Else
                CurrentEXPBarWidth = CurrentEXPBarWidth + MoveSpeed
            End If

        Else
            MoveSpeed = (OldEXPBarWidth - NewEXPBarWidth) / 25

            If MoveSpeed < 1 Then MoveSpeed = 1

            If CurrentEXPBarWidth + MoveSpeed < NewEXPBarWidth Then
                CurrentEXPBarWidth = NewEXPBarWidth
            Else
                CurrentEXPBarWidth = CurrentEXPBarWidth - MoveSpeed
            End If
            
        End If
        
        If CurrentEXPBarWidth >= 0 Then frmMain.imgEXPBar.Width = CurrentEXPBarWidth
    End If
    Exit Sub
    
' Error Handler
ErrorHandler:
    HandleError "ResizeMPBar", "modGeneral", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawEquipment()
    Dim I As Long
    Dim ItemNum As Long
    Dim ItemPic As Long
    Dim sRECT As RECT
    Dim dRect As RECT
    Dim destPresentationRect As D3DRECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    ' Render equipment base
    With sRECT
        .Top = 0
        .Bottom = Tex_Equip.Height
        .Left = 0
        .Right = Tex_Equip.Width
    End With
    
    With dRect
        .Top = 0
        .Bottom = frmMain.picEquipment.Height
        .Left = 0
        .Right = frmMain.picEquipment.Width
    End With

    RenderTextureByRects Tex_Equip, sRECT, dRect
    
    ' Now lets make the image that we will be rendering today
    For I = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(MyIndex, I)
        
        ' If there is an item draw it, if not do NOTHING!
        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            ItemPic = Item(ItemNum).Pic
            
            ' If the picture exists then render it
            If ItemPic > 0 And ItemPic <= NumItems Then
                sRECT.Top = 0
                sRECT.Bottom = PIC_Y
                sRECT.Left = 0
                sRECT.Right = PIC_X

                RenderTexture Tex_Item(ItemPic), EquipSlotLeft(I), EquipSlotTop(I), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top
            End If
        End If
    Next
    
    With destPresentationRect
        .X1 = 0
        .X2 = Tex_Equip.Width
        .Y1 = 0
        .Y2 = Tex_Equip.Height
    End With
    
    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destPresentationRect, destPresentationRect, frmMain.picEquipment.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawEquipment", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorEmoticon_DrawIcon()
    Dim EmoticonNum As Long
    Dim sRECT As RECT
    Dim dRect As RECT
    Dim destRECT As D3DRECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    EmoticonNum = frmEditor_Emoticon.scrlEmoticon.Value

    If EmoticonNum < 1 Or EmoticonNum > NumEmoticons Then
        frmEditor_Emoticon.picEmoticon.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = Tex_Emoticon(EmoticonNum).Height
    sRECT.Left = 0
    sRECT.Right = sRECT.Left + Tex_Emoticon(EmoticonNum).Width
    dRect.Top = 0
    dRect.Bottom = Tex_Emoticon(EmoticonNum).Height
    dRect.Left = 0
    dRect.Right = Tex_Emoticon(EmoticonNum).Width
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Emoticon(EmoticonNum), sRECT, dRect
    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_Emoticon.picEmoticon.Width
        .Y1 = 0
        .Y2 = frmEditor_Emoticon.picEmoticon.Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Emoticon.picEmoticon.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorEmoticon_BltIcon", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawEmoticons()
    Dim sRECT As RECT
    Dim dRect As RECT
    Dim EmoticonNum As Byte, I As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerMap(MyIndex) = GetPlayerMap(I) Then
            EmoticonNum = TempPlayer(I).EmoticonNum
            
            If EmoticonNum < 1 Or EmoticonNum > NumEmoticons Then
                If Trim$(Player(MyIndex).Status) = "AFK" Then
                    EmoticonNum = Emoticon(1).Pic
                Else
                    Exit Sub
                End If
            End If
            
            ' Clear out the data if it needs to disappear
            If timeGetTime > TempPlayer(I).EmoticonTimer And EmoticonNum <> Emoticon(1).Pic Then
                TempPlayer(I).EmoticonNum = 0
                TempPlayer(I).EmoticonTimer = 0
                Exit Sub
            End If
    
            If InViewPort(GetPlayerX(I), GetPlayerY(I)) Then
                With sRECT
                    .Top = 0
                    .Bottom = .Top + PIC_Y
                    .Left = 0
                    .Right = .Left + PIC_X
                End With
                
                ' Same for destination as source
                dRect = sRECT
                
                RenderTexture Tex_Emoticon(EmoticonNum), GetPlayerTextX(I) - 16, GetPlayerTextY(I) - 16, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawEmoticons", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorAnim_DrawAnim()
    Dim I As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, Height As Long, looptime As Long, FrameCount As Long
    Dim sX As Long, sY As Long, sRECT As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    For I = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(I).Value
        
        If Animationnum <= 0 Or Animationnum > NumAnimations Then
            frmEditor_Animation.picSprite.Item(I).Cls
        Else
            sRECT.Top = 0
            sRECT.Left = 0
            sRECT.Right = Tex_Animation(Animationnum).Width \ AnimColumns
            sRECT.Bottom = sRECT.Right
        
            looptime = frmEditor_Animation.scrlLoopTime(I)
            FrameCount = frmEditor_Animation.scrlFrameCount(I)
            
            ShouldRender = False
            
            ' Check if we need to render new frame
            If AnimEditorTimer(I) + looptime <= timeGetTime Then
                ' Check if out of range
                If AnimEditorFrame(I) >= FrameCount Then
                    AnimEditorFrame(I) = 1
                Else
                    AnimEditorFrame(I) = AnimEditorFrame(I) + 1
                End If
                AnimEditorTimer(I) = timeGetTime
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(I).Value > 0 Then
                    ' Total width divided by frame count
                    Width = Tex_Animation(Animationnum).Width \ AnimColumns
                    Height = Width

                    sY = (Height * ((AnimEditorFrame(I) - 1) \ AnimColumns))
                    sX = (Width * (((AnimEditorFrame(I) - 1) Mod AnimColumns)))

                    ' Start Rendering
                    Call Direct3D_Device.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call Direct3D_Device.BeginScene
                    
                    RenderTexture Tex_Animation(Animationnum), 0, 0, sX, sY, Width, Height, Width, Height
                    
                    ' Finish Rendering
                    Call Direct3D_Device.EndScene
                    Call Direct3D_Device.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(I).hWnd, ByVal 0)
                End If
            End If
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorAnim_DrawAnim", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorNPC_DrawSprite()
    Dim Sprite As Long, destRECT As D3DRECT
    Dim sRECT As RECT
    Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Sprite = frmEditor_NPC.scrlSprite.Value

    If Sprite < 1 Or Sprite > NumCharacters Then
        frmEditor_NPC.picSprite.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = Tex_Character(Sprite).Height / 4
    sRECT.Left = (Tex_Character(Sprite).Width / 4) * 2 ' Facing down
    sRECT.Right = sRECT.Left + Tex_Character(Sprite).Width / 4
    dRect.Top = 0
    dRect.Bottom = Tex_Character(Sprite).Height / 4
    dRect.Left = 0
    dRect.Right = Tex_Character(Sprite).Width / 4
    
    frmEditor_NPC.picSprite.Width = Tex_Character(Sprite).Width / 4
    frmEditor_NPC.picSprite.Height = Tex_Character(Sprite).Height / 4
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Character(Sprite), sRECT, dRect
    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_NPC.picSprite.Width
        .Y1 = 0
        .Y2 = frmEditor_NPC.picSprite.Height
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_NPC.picSprite.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorNPC_DrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorResource_DrawSprite()
    Dim Sprite As Long
    Dim sRECT As RECT, destRECT As D3DRECT, srcRect As D3DRECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Normal sprite
    Sprite = frmEditor_Resource.scrlNormalPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picNormalPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(Sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRECT, dRect
        
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
        
        With destRECT
            .X1 = 0
            .X2 = frmEditor_Resource.picNormalPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picNormalPic.ScaleHeight
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmEditor_Resource.picNormalPic.hWnd, ByVal (0)
    End If

    ' Exhausted sprite
    Sprite = frmEditor_Resource.scrlExhaustedPic.Value

    If Sprite < 1 Or Sprite > NumResources Then
        frmEditor_Resource.picExhaustedPic.Cls
    Else
        sRECT.Top = 0
        sRECT.Bottom = Tex_Resource(Sprite).Height
        sRECT.Left = 0
        sRECT.Right = Tex_Resource(Sprite).Width
        dRect.Top = 0
        dRect.Bottom = Tex_Resource(Sprite).Height
        dRect.Left = 0
        dRect.Right = Tex_Resource(Sprite).Width
        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
        Direct3D_Device.BeginScene
        RenderTextureByRects Tex_Resource(Sprite), sRECT, dRect
        
        With destRECT
            .X1 = 0
            .X2 = frmEditor_Resource.picExhaustedPic.ScaleWidth
            .Y1 = 0
            .Y2 = frmEditor_Resource.picExhaustedPic.ScaleHeight
        End With
        
        With srcRect
            .X1 = 0
            .X2 = Tex_Resource(Sprite).Width
            .Y1 = 0
            .Y2 = Tex_Resource(Sprite).Height
        End With
                    
        Direct3D_Device.EndScene
        Direct3D_Device.Present srcRect, destRECT, frmEditor_Resource.picExhaustedPic.hWnd, ByVal (0)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorResource_DrawSprite", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorMap_DrawMapItem()
    Dim ItemNum As Long
    Dim sRECT As RECT, destRECT As D3DRECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ItemNum = Item(frmEditor_Map.scrlMapItem.Value).Pic

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmEditor_Map.picMapItem.Cls
        Exit Sub
    End If

    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(ItemNum), sRECT, dRect
    
    With destRECT
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Map.picMapItem.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorMap_DrawMapItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorItem_DrawItem()
    Dim ItemNum As Long
    Dim sRECT As RECT, destRECT As D3DRECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ItemNum = frmEditor_Item.scrlPic.Value

    If ItemNum < 1 Or ItemNum > NumItems Then
        frmEditor_Item.picItem.Cls
        Exit Sub
    End If

    ' Rect for source
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    
    ' Same for destination as source
    dRect = sRECT
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Item(ItemNum), sRECT, dRect
    
    With destRECT
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Item.picItem.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorItem_DrawItem", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorItem_DrawPaperdoll()
    Dim Sprite As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim sRECT As RECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    Sprite = frmEditor_Item.scrlPaperdoll.Value

    If Sprite < 1 Or Sprite > NumPaperdolls Then
        frmEditor_Item.picPaperdoll.Cls
        Exit Sub
    End If

    ' Rect for source
    sRECT.Top = 0
    sRECT.Bottom = Tex_Paperdoll(Sprite).Height / 4
    sRECT.Left = 0
    sRECT.Right = Tex_Paperdoll(Sprite).Width / 4
    
    ' Same for destination as source
    dRect = sRECT
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_Paperdoll(Sprite), sRECT, dRect
                    
    With destRECT
        .X1 = 0
        .X2 = Tex_Paperdoll(Sprite).Width / 4
        .Y1 = 0
        .Y2 = Tex_Paperdoll(Sprite).Height / 4
    End With
                    
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Item.picPaperdoll.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorItem_DrawPaperdoll", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorSpell_DrawIcon()
    Dim IconNum As Long, destRECT As D3DRECT
    Dim sRECT As RECT
    Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    IconNum = frmEditor_Spell.scrlIcon.Value
    
    If IconNum < 1 Or IconNum > NumSpellIcons Then
        frmEditor_Spell.picSprite.Cls
        Exit Sub
    End If
    
    sRECT.Top = 0
    sRECT.Bottom = PIC_Y
    sRECT.Left = 0
    sRECT.Right = PIC_X
    dRect.Top = 0
    dRect.Bottom = PIC_Y
    dRect.Left = 0
    dRect.Right = PIC_X
    
    With destRECT
        .X1 = 0
        .X2 = PIC_X
        .Y1 = 0
        .Y2 = PIC_Y
    End With
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    RenderTextureByRects Tex_SpellIcon(IconNum), sRECT, dRect
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_Spell.picSprite.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorSpell_DrawIcon", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorMap_DrawTileset()
    Dim Height As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Width As Long
    Dim Tileset As Long
    Dim sRECT As RECT
    Dim dRect As RECT, scrlX As Long, scrlY As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Find tileset number
    Tileset = frmEditor_Map.scrlTileSet.Value
    
    ' Exit out if doesn't exist
    If Tileset < 0 Or Tileset > NumTileSets Then Exit Sub

    Height = Tex_Tileset(Tileset).Height
    Width = Tex_Tileset(Tileset).Width
    
    RenderTexture Tex_Fade, 0, 0, 0, 0, ScreenX, Tex_Tileset(Tileset).Height + 32, 32, 32
    
    RenderTexture Tex_Tileset(Tileset), 0, 0, 0, 0, Tex_Tileset(Tileset).Width, Tex_Tileset(Tileset).Height, Tex_Tileset(Tileset).Width, Tex_Tileset(Tileset).Height
    
    With dRect
        .Bottom = 0
        .Left = (Tex_Tileset(Tileset).Width - 115 - scrlX) / 2
        .Right = 0
        .Top = Tex_Tileset(Tileset).Height + 10 - scrlY
    End With
    
    With destRECT
        .X1 = (EditorTileX * 32) - sRECT.Left
        .X2 = (EditorTileWidth * 32) + .X1
        .Y1 = (EditorTileY * 32) - sRECT.Top
        .Y2 = (EditorTileHeight * 32) + .Y1
    End With
    
    DrawSelectionBox destRECT
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorMap_DrawTileset", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub DrawSelectionBox(dRect As D3DRECT)
    Dim Width As Long, Height As Long, X As Long, Y As Long
    
    Width = dRect.X2 - dRect.X1
    Height = dRect.Y2 - dRect.Y1
    X = dRect.X1
    Y = dRect.Y1
    
    If Width > 6 And Height > 6 Then
        ' Draw Box 32 by 32 at graphicselx and graphicsely
        RenderTexture Tex_Selection, X, Y, 1, 1, 2, 2, 2, 2, -1 ' Top left corner
        RenderTexture Tex_Selection, X + 2, Y, 3, 1, Width - 4, 2, 32 - 6, 2, -1 ' Top line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y, 29, 1, 2, 2, 2, 2, -1 ' Top right corner
        RenderTexture Tex_Selection, X, Y + 2, 1, 3, 2, Height - 4, 2, 32 - 6, -1 ' Left Line
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2, 32 - 3, 3, 2, Height - 4, 2, 32 - 6, -1 ' Right line
        RenderTexture Tex_Selection, X, Y + 2 + (Height - 4), 1, 32 - 3, 2, 2, 2, 2, -1 ' Bottom left corner
        RenderTexture Tex_Selection, X + 2 + (Width - 4), Y + 2 + (Height - 4), 32 - 3, 32 - 3, 2, 2, 2, 2, -1 ' Bottom right corner
        RenderTexture Tex_Selection, X + 2, Y + 2 + (Height - 4), 3, 32 - 3, Width - 4, 2, 32 - 6, 2, -1 ' Bottom line
    End If
End Sub

Public Sub DrawEvents()
    Dim sRECT As RECT
    Dim Width As Long, Height As Long, I As Long, X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Map.EventCount <= 0 Then Exit Sub
    
    For I = 1 To Map.EventCount
        If Map.events(I).PageCount <= 0 Then
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                RenderTexture Tex_Selection, ConvertMapX(X), ConvertMapY(Y), sRECT.Left, sRECT.Right, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            GoTo nextevent
        End If
        
        Width = 32
        Height = 32
    
        X = Map.events(I).X * 32
        Y = Map.events(I).Y * 32
        X = ConvertMapX(X)
        Y = ConvertMapY(Y)
    
        If I > Map.EventCount Then Exit Sub
        If 1 > Map.events(I).PageCount Then Exit Sub
        
        Select Case Map.events(I).Pages(1).GraphicType
            Case 0
                sRECT.Top = 0
                sRECT.Bottom = 32
                sRECT.Left = 0
                sRECT.Right = 32
                RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            Case 1
                If Map.events(I).Pages(1).Graphic > 0 And Map.events(I).Pages(1).Graphic <= NumCharacters Then
                    
                    sRECT.Top = (Map.events(I).Pages(1).GraphicY * (Tex_Character(Map.events(I).Pages(1).Graphic).Height / 4))
                    sRECT.Left = (Map.events(I).Pages(1).GraphicX * (Tex_Character(Map.events(I).Pages(1).Graphic).Width / 4))
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    RenderTexture Tex_Character(Map.events(I).Pages(1).Graphic), X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
            Case 2
                If Map.events(I).Pages(1).Graphic > 0 And Map.events(I).Pages(1).Graphic < NumTileSets Then
                    sRECT.Top = Map.events(I).Pages(1).GraphicY * 32
                    sRECT.Left = Map.events(I).Pages(1).GraphicX * 32
                    sRECT.Bottom = sRECT.Top + 32
                    sRECT.Right = sRECT.Left + 32
                    RenderTexture Tex_Tileset(Map.events(I).Pages(1).Graphic), X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                    
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                Else
                    sRECT.Top = 0
                    sRECT.Bottom = 32
                    sRECT.Left = 0
                    sRECT.Right = 32
                    RenderTexture Tex_Selection, X, Y, sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
                End If
        End Select
        
nextevent:
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "DrawEvents", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub EditorEvent_DrawGraphic()
    Dim sRECT As RECT, destRECT As D3DRECT, srcRect As D3DRECT
    Dim dRect As RECT

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If frmEditor_Events.picGraphicSel.Visible Then
        Select Case frmEditor_Events.cmbGraphic.ListIndex
            Case 0
                ' None
                frmEditor_Events.picGraphicSel.Cls
                Exit Sub
            Case 1
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumCharacters Then
                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Right = sRECT.Left + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width - sRECT.Left)
                    Else
                        sRECT.Left = 0
                        sRECT.Right = Tex_Character(frmEditor_Events.scrlGraphic.Value).Width
                    End If
                    
                    If Tex_Character(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRECT.Top = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Bottom = sRECT.Top + (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height - sRECT.Top)
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = Tex_Character(frmEditor_Events.scrlGraphic.Value).Height
                    End If
                    
                    With dRect
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    With destRECT
                        .X1 = dRect.Left
                        .X2 = dRect.Right
                        .Y1 = dRect.Top
                        .Y2 = dRect.Bottom
                    End With
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRECT, dRect
                    
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRECT
                            .X1 = (GraphicSelX * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4)) - sRECT.Left
                            .X2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Width / 4) + .X1
                            .Y1 = (GraphicSelY * (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4)) - sRECT.Top
                            .Y2 = (Tex_Character(frmEditor_Events.scrlGraphic.Value).Height / 4) + .Y1
                        End With

                    Else
                        With destRECT
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    DrawSelectionBox destRECT
                    
                    With srcRect
                        .X1 = dRect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = dRect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRECT
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRECT, frmEditor_Events.picGraphicSel.hWnd, ByVal (0)
                    
                    If GraphicSelX <= 3 And GraphicSelY <= 3 Then
                    Else
                        GraphicSelX = 0
                        GraphicSelY = 0
                    End If
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
            Case 2
                If frmEditor_Events.scrlGraphic.Value > 0 And frmEditor_Events.scrlGraphic.Value <= NumTileSets Then
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width > 793 Then
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value
                        sRECT.Right = sRECT.Left + 800
                    Else
                        sRECT.Left = 0
                        sRECT.Right = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Width
                        sRECT.Left = frmEditor_Events.hScrlGraphicSel.Value = 0
                    End If
                    
                    If Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height > 472 Then
                        sRECT.Top = frmEditor_Events.vScrlGraphicSel.Value
                        sRECT.Bottom = sRECT.Top + 512
                    Else
                        sRECT.Top = 0
                        sRECT.Bottom = Tex_Tileset(frmEditor_Events.scrlGraphic.Value).Height
                        frmEditor_Events.vScrlGraphicSel.Value = 0
                    End If
                    
                    If sRECT.Left = -1 Then sRECT.Left = 0
                    If sRECT.Top = -1 Then sRECT.Top = 0
                    
                    With dRect
                        .Top = 0
                        .Bottom = sRECT.Bottom - sRECT.Top
                        .Left = 0
                        .Right = sRECT.Right - sRECT.Left
                    End With
                    
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRect
                    
                    If (GraphicSelX2 < GraphicSelX Or GraphicSelY2 < GraphicSelY) Or (GraphicSelX2 = 0 And GraphicSelY2 = 0) Then
                        With destRECT
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = PIC_X + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = PIC_Y + .Y1
                        End With

                    Else
                        With destRECT
                            .X1 = (GraphicSelX * 32) - sRECT.Left
                            .X2 = ((GraphicSelX2 - GraphicSelX) * 32) + .X1
                            .Y1 = (GraphicSelY * 32) - sRECT.Top
                            .Y2 = ((GraphicSelY2 - GraphicSelY) * 32) + .Y1
                        End With
                    End If
                    
                    DrawSelectionBox destRECT
                    
                    With srcRect
                        .X1 = dRect.Left
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y1 = dRect.Top
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    With destRECT
                        .X1 = 0
                        .Y1 = 0
                        .X2 = frmEditor_Events.picGraphicSel.ScaleWidth
                        .Y2 = frmEditor_Events.picGraphicSel.ScaleHeight
                    End With
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present srcRect, destRECT, frmEditor_Events.picGraphicSel.hWnd, ByVal (0)
                Else
                    frmEditor_Events.picGraphicSel.Cls
                    Exit Sub
                End If
        End Select
    Else
        If curPageNum < 1 Then Exit Sub ' Without this line it was Unrecoverable DX8 Error
        Select Case tmpEvent.Pages(curPageNum).GraphicType
            Case 0
                frmEditor_Events.picGraphic.Cls
            Case 1
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumCharacters Then
                    sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    sRECT.Bottom = sRECT.Top + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Height / 4)
                    sRECT.Right = sRECT.Left + (Tex_Character(tmpEvent.Pages(curPageNum).Graphic).Width / 4)
                    With dRect
                        dRect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                        dRect.Bottom = dRect.Top + (sRECT.Bottom - sRECT.Top)
                        dRect.Left = (121 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                        dRect.Right = dRect.Left + (sRECT.Right - sRECT.Left)
                    End With
                    With destRECT
                        .X1 = dRect.Left
                        .X2 = dRect.Right
                        .Y1 = dRect.Top
                        .Y2 = dRect.Bottom
                    End With
                    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                    Direct3D_Device.BeginScene
                    RenderTextureByRects Tex_Character(frmEditor_Events.scrlGraphic.Value), sRECT, dRect
                    Direct3D_Device.EndScene
                    Direct3D_Device.Present destRECT, destRECT, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                End If
            Case 2
                If tmpEvent.Pages(curPageNum).Graphic > 0 And tmpEvent.Pages(curPageNum).Graphic <= NumTileSets Then
                    If tmpEvent.Pages(curPageNum).GraphicX2 = 0 Or tmpEvent.Pages(curPageNum).GraphicY2 = 0 Then
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + 32
                        sRECT.Right = sRECT.Left + 32
                        With dRect
                            dRect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRECT.Bottom - sRECT.Top)
                            dRect.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            dRect.Right = dRect.Left + (sRECT.Right - sRECT.Left)
                        End With
                        With destRECT
                            .X1 = dRect.Left
                            .X2 = dRect.Right
                            .Y1 = dRect.Top
                            .Y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRECT, destRECT, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                    Else
                        sRECT.Top = tmpEvent.Pages(curPageNum).GraphicY * 32
                        sRECT.Left = tmpEvent.Pages(curPageNum).GraphicX * 32
                        sRECT.Bottom = sRECT.Top + ((tmpEvent.Pages(curPageNum).GraphicY2 - tmpEvent.Pages(curPageNum).GraphicY) * 32)
                        sRECT.Right = sRECT.Left + ((tmpEvent.Pages(curPageNum).GraphicX2 - tmpEvent.Pages(curPageNum).GraphicX) * 32)
                        With dRect
                            dRect.Top = (193 / 2) - ((sRECT.Bottom - sRECT.Top) / 2)
                            dRect.Bottom = dRect.Top + (sRECT.Bottom - sRECT.Top)
                            dRect.Left = (120 / 2) - ((sRECT.Right - sRECT.Left) / 2)
                            dRect.Right = dRect.Left + (sRECT.Right - sRECT.Left)
                        End With
                        With destRECT
                            .X1 = dRect.Left
                            .X2 = dRect.Right
                            .Y1 = dRect.Top
                            .Y2 = dRect.Bottom
                        End With
                        Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
                        Direct3D_Device.BeginScene
                        RenderTextureByRects Tex_Tileset(frmEditor_Events.scrlGraphic.Value), sRECT, dRect
                        Direct3D_Device.EndScene
                        Direct3D_Device.Present destRECT, destRECT, frmEditor_Events.picGraphic.hWnd, ByVal (0)
                    End If
                End If
        End Select
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorMap_DrawKey", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DrawEvent(ID As Long)
    Dim X As Long, Y As Long, Width As Long, Height As Long, sRECT As RECT, dRect As RECT, Anim As Long, spritetop As Long
    
    If Map.MapEvents(ID).Visible = 0 Then Exit Sub
    If InMapEditor Then Exit Sub
    
    Select Case Map.MapEvents(ID).GraphicType
        Case 0
            Exit Sub
        Case 1
            If Map.MapEvents(ID).GraphicNum <= 0 Or Map.MapEvents(ID).GraphicNum > NumCharacters Then Exit Sub
            Width = Tex_Character(Map.MapEvents(ID).GraphicNum).Width / 4
            Height = Tex_Character(Map.MapEvents(ID).GraphicNum).Height / 4
            
            If Map.MapEvents(ID).Moving > 0 Then
                Anim = Map.MapEvents(ID).Step
            Else
                Anim = 0
            End If
            
            ' Set the left
            Select Case Map.MapEvents(ID).ShowDir
                Case DIR_UP
                    spritetop = 3
                Case DIR_RIGHT
                    spritetop = 2
                Case DIR_DOWN
                    spritetop = 0
                Case DIR_LEFT
                    spritetop = 1
            End Select
            
            If Map.MapEvents(ID).WalkAnim = 1 Then Anim = 0
            
            If Map.MapEvents(ID).Moving = 0 Then Anim = Map.MapEvents(ID).GraphicX
            
            With sRECT
                .Top = spritetop * Height
                .Bottom = .Top + Height
                .Left = Anim * Width
                .Right = .Left + Width
            End With
        
            ' Calculate the X
            X = Map.MapEvents(ID).X * PIC_X + Map.MapEvents(ID).xOffset - ((Width - 32) / 2)
        
            ' Is the player's height more than 32..?
            If (Height * 4) > 32 Then
                ' Create a 32 pixel offset for larger sprites
                Y = Map.MapEvents(ID).Y * PIC_Y + Map.MapEvents(ID).yOffset - ((Height) - 32)
            Else
                ' Proceed as normal
                Y = Map.MapEvents(ID).Y * PIC_Y + Map.MapEvents(ID).yOffset
            End If
        
            ' render the actual sprite
            Call DrawSprite(Map.MapEvents(ID).GraphicNum, X, Y, sRECT)
            
        Case 2
            If Map.MapEvents(ID).GraphicNum < 1 Or Map.MapEvents(ID).GraphicNum > NumTileSets Then Exit Sub
            
            If Map.MapEvents(ID).GraphicY2 > 0 Or Map.MapEvents(ID).GraphicX2 > 0 Then
                With sRECT
                    .Top = Map.MapEvents(ID).GraphicY * 32
                    .Bottom = .Top + ((Map.MapEvents(ID).GraphicY2 - Map.MapEvents(ID).GraphicY) * 32)
                    .Left = Map.MapEvents(ID).GraphicX * 32
                    .Right = .Left + ((Map.MapEvents(ID).GraphicX2 - Map.MapEvents(ID).GraphicX) * 32)
                End With
            Else
                With sRECT
                    .Top = Map.MapEvents(ID).GraphicY * 32
                    .Bottom = .Top + 32
                    .Left = Map.MapEvents(ID).GraphicX * 32
                    .Right = .Left + 32
                End With
            End If
            
            X = Map.MapEvents(ID).X * 32
            Y = Map.MapEvents(ID).Y * 32
            
            X = X - ((sRECT.Right - sRECT.Left) / 2)
            Y = Y - (sRECT.Bottom - sRECT.Top) + 32
            
            
            If Map.MapEvents(ID).GraphicY2 > 0 Then
                RenderTexture Tex_Tileset(Map.MapEvents(ID).GraphicNum), ConvertMapX(Map.MapEvents(ID).X * 32), ConvertMapY((Map.MapEvents(ID).Y - ((Map.MapEvents(ID).GraphicY2 - Map.MapEvents(ID).GraphicY) - 1)) * 32), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            Else
                RenderTexture Tex_Tileset(Map.MapEvents(ID).GraphicNum), ConvertMapX(Map.MapEvents(ID).X * 32), ConvertMapY(Map.MapEvents(ID).Y * 32), sRECT.Left, sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, sRECT.Right - sRECT.Left, sRECT.Bottom - sRECT.Top, D3DColorRGBA(255, 255, 255, 255)
            End If
    End Select
End Sub

Sub HandleDeviceLost()
    On Error Resume Next
    
    ' Do a loop while device is lost
    Do While Direct3D_Device.TestCooperativeLevel = D3DERR_DEVICELOST
        Exit Sub
    Loop
    
    UnloadTextures
    
    ' Reset the device
    Direct3D_Device.Reset Direct3D_Window
    
    DirectX_ReInit
     
    LoadTextures
End Sub

Private Function DirectX_ReInit() As Boolean
    On Error GoTo Error_Handler

    Direct3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, Display_Mode 'Use the current display mode that you
                                                                    'are already on. Incase you are confused, I'm
                                                                    'talking about your current screen resolution. ;)
        
    Direct3D_Window.Windowed = True 'The app will be in windowed mode.

    Direct3D_Window.SwapEffect = D3DSWAPEFFECT_COPY 'Refresh when the monitor does.
    Direct3D_Window.BackBufferFormat = Display_Mode.Format 'Sets the format that was retrieved into the backbuffer.
    'Creates the rendering device with some useful info, along with the info
    'we've already setup for Direct3D_Window.
    'Creates the rendering device with some useful info, along with the info
    Direct3D_Window.BackBufferCount = 1 '1 backbuffer only
    Direct3D_Window.BackBufferWidth = 800 ' FrmMain.picScreen.ScaleWidth 'Match the backbuffer width with the display width
    Direct3D_Window.BackBufferHeight = 600 'frmMain.picScreen.ScaleHeight 'Match the backbuffer height with the display height
    Direct3D_Window.hDeviceWindow = frmMain.picScreen.hWnd 'Use frmMain as the device window.
    
    With Direct3D_Device
        .SetVertexShader D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetRenderState D3DRS_ZENABLE, False
        .SetRenderState D3DRS_ZWRITEENABLE, False
        
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
        .SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
        .SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    
        .SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_POINT
        .SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_POINT
    End With
    
    DirectX_ReInit = True

    Exit Function
    
Error_Handler:
    MsgBox "An error occured while initializing DirectX", vbCritical
    
    DestroyGame
    
    DirectX_ReInit = False
End Function

Public Sub UpdateCamera()

    Dim offsetX As Long

    Dim offsetY As Long

    Dim StartX As Long

    Dim StartY As Long

    Dim EndX As Long

    Dim EndY As Long

    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    offsetX = TempPlayer(MyIndex).xOffset + PIC_X
    offsetY = TempPlayer(MyIndex).yOffset + PIC_Y
    StartX = GetPlayerX(MyIndex) - ((MIN_MAPX + 1) \ 2) - 1
    StartY = GetPlayerY(MyIndex) - ((MIN_MAPY + 1) \ 2) - 1

    If StartX < 0 Then
        offsetX = 0

        If StartX = -1 Then
            If TempPlayer(MyIndex).xOffset > 0 Then
                offsetX = TempPlayer(MyIndex).xOffset
            End If
        End If

        StartX = 0
    End If

    If StartY < 0 Then
        offsetY = 0

        If StartY = -1 Then
            If TempPlayer(MyIndex).yOffset > 0 Then
                offsetY = TempPlayer(MyIndex).yOffset
            End If
        End If

        StartY = 0
    End If

    EndX = StartX + (MIN_MAPX + 1)
    EndY = StartY + (MIN_MAPY + 1)

    If EndX >= Map.MaxX Then
        offsetX = 32
        
        If EndX = Map.MaxX Then
            If TempPlayer(MyIndex).xOffset < 0 Then
                offsetX = TempPlayer(MyIndex).xOffset + PIC_X
            End If
        End If

        EndX = Map.MaxX
        StartX = EndX - MIN_MAPX - 1
    End If

    If EndY >= Map.MaxY Then
        offsetY = 32

        If EndY = Map.MaxY Then
            If TempPlayer(MyIndex).yOffset < 0 Then
                offsetY = TempPlayer(MyIndex).yOffset + PIC_Y
            End If
        End If

        EndY = Map.MaxY
        StartY = EndY - MIN_MAPY - 1
    End If

    With TileView
        .Top = StartY
        .Bottom = EndY
        .Left = StartX
        .Right = EndX
    End With

    With Camera
        .Top = offsetY
        .Bottom = .Top + ScreenY
        .Left = offsetX
        .Right = .Left + ScreenX
    End With

    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "UpdateCamera", "modGraphics", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Public Sub InitAutotiles()
    Dim X As Long, Y As Long, layerNum As Long
    
    ' Procedure used to cache autotile positions. All positioning is
    ' Independant from the tileset. Calculations are convoluted and annoying.
    ' Maths is not my strong point. Luckily we're caching them so it's a one-off
    ' thing when the map is originally loaded. As such optimisation isn't an issue.
    
    ' For simplicity's sake we cache all subtile SOURCE positions in to an array.
    ' We also give letters to each subtile for easy rendering tweaks. ;]
    
    ' First, we need to re-size the array
    ReDim Autotile(0 To Map.MaxX, 0 To Map.MaxY)
    
    ' Inner tiles (Top right subtile region)
    ' NW - a
    autoInner(1).X = 32
    autoInner(1).Y = 0
    
    ' NE - b
    autoInner(2).X = 48
    autoInner(2).Y = 0
    
    ' SW - c
    autoInner(3).X = 32
    autoInner(3).Y = 16
    
    ' SE - d
    autoInner(4).X = 48
    autoInner(4).Y = 16
    
    ' Outer Tiles - NW (bottom subtile region)
    ' NW - e
    autoNW(1).X = 0
    autoNW(1).Y = 32
    
    ' NE - f
    autoNW(2).X = 16
    autoNW(2).Y = 32
    
    ' SW - g
    autoNW(3).X = 0
    autoNW(3).Y = 48
    
    ' SE - h
    autoNW(4).X = 16
    autoNW(4).Y = 48
    
    ' Outer Tiles - NE (bottom subtile region)
    ' NW - i
    autoNE(1).X = 32
    autoNE(1).Y = 32
    
    ' NE - g
    autoNE(2).X = 48
    autoNE(2).Y = 32
    
    ' SW - k
    autoNE(3).X = 32
    autoNE(3).Y = 48
    
    ' SE - l
    autoNE(4).X = 48
    autoNE(4).Y = 48
    
    ' Outer Tiles - SW (bottom subtile region)
    ' NW - m
    autoSW(1).X = 0
    autoSW(1).Y = 64
    
    ' NE - n
    autoSW(2).X = 16
    autoSW(2).Y = 64
    
    ' SW - o
    autoSW(3).X = 0
    autoSW(3).Y = 80
    
    ' SE - p
    autoSW(4).X = 16
    autoSW(4).Y = 80
    
    ' Outer Tiles - SE (bottom subtile region)
    ' NW - q
    autoSE(1).X = 32
    autoSE(1).Y = 64
    
    ' NE - r
    autoSE(2).X = 48
    autoSE(2).Y = 64
    
    ' SW - s
    autoSE(3).X = 32
    autoSE(3).Y = 80
    
    ' SE - t
    autoSE(4).X = 48
    autoSE(4).Y = 80
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For layerNum = 1 To MapLayer.Layer_Count - 1
                ' Calculate the subtile positions and place them
                CalculateAutotile X, Y, layerNum
                
                ' Cache the rendering state of the tiles and set them
                CacheRenderState X, Y, layerNum
            Next
        Next
    Next
End Sub

Public Sub CacheRenderState(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    Dim quarterNum As Long

    ' Exit out early
    If X < 0 Or X > Map.MaxX Or Y < 0 Or Y > Map.MaxY Then Exit Sub

    With Map.Tile(X, Y)
        ' Check if the tile can be rendered
        If .Layer(layerNum).Tileset <= 0 Or .Layer(layerNum).Tileset > NumTileSets Then
            Autotile(X, Y).Layer(layerNum).RenderState = RENDER_STATE_NONE
            Exit Sub
        End If
        
        ' Check if it needs to be rendered as an autotile
        If .Autotile(layerNum) = AUTOTILE_NONE Or .Autotile(layerNum) = AUTOTILE_FAKE Then
            ' default to... default
            Autotile(X, Y).Layer(layerNum).RenderState = RENDER_STATE_NORMAL
        Else
            Autotile(X, Y).Layer(layerNum).RenderState = RENDER_STATE_AUTOTILE
            ' Cache tileset positioning
            For quarterNum = 1 To 4
                Autotile(X, Y).Layer(layerNum).srcX(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).X * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).X
                Autotile(X, Y).Layer(layerNum).srcY(quarterNum) = (Map.Tile(X, Y).Layer(layerNum).Y * 32) + Autotile(X, Y).Layer(layerNum).QuarterTile(quarterNum).Y
            Next
        End If
    End With
End Sub

Public Sub CalculateAutotile(ByVal X As Long, ByVal Y As Long, ByVal layerNum As Long)
    ' Right, so we've split the tile block in to an easy to remember
    ' Collection of letters. We now need to do the calculations to find
    ' out which little lettered block needs to be rendered. We do this
    ' by reading the surrounding tiles to check for matches.
    
    ' First we check to make sure an autotile situation is actually there.
    ' Then we calculate exactly which situation has arisen.
    ' The situations are "inner", "outer", "horizontal", "vertical" and "fill".
    
    ' Exit out if we don't have an auatotile
    If Map.Tile(X, Y).Autotile(layerNum) = 0 Then Exit Sub
    
    ' Okay, we have autotiling but which one?
    Select Case Map.Tile(X, Y).Autotile(layerNum)
        ' Normal or animated - same difference
        Case AUTOTILE_NORMAL, AUTOTILE_ANIM
            ' North West Quarter
            CalculateNW_Normal layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Normal layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Normal layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Normal layerNum, X, Y
            
        ' Cliff
        Case AUTOTILE_CLIFF
            ' North West Quarter
            CalculateNW_Cliff layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Cliff layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Cliff layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Cliff layerNum, X, Y
            
        ' Waterfalls
        Case AUTOTILE_WATERFALL
            ' North West Quarter
            CalculateNW_Waterfall layerNum, X, Y
            
            ' North East Quarter
            CalculateNE_Waterfall layerNum, X, Y
            
            ' South West Quarter
            CalculateSW_Waterfall layerNum, X, Y
            
            ' South East Quarter
            CalculateSE_Waterfall layerNum, X, Y
        
        ' Anything else
        Case Else
            ' Don't need to render anything... it's fake or not an autotile
    End Select
End Sub

' Normal autotiling
Public Sub CalculateNW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If Not tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 1, "a"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 2, "b"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 3, "c"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Normal(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    ' Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Outer
    If tmpTile(1) And Not tmpTile(2) And tmpTile(3) Then situation = AUTO_OUTER
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_OUTER
            placeAutotile layerNum, X, Y, 4, "d"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

' Waterfall autotiling
Public Sub CalculateNW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 1, "i"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 1, "e"
    End If
End Sub

Public Sub CalculateNE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 2, "f"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 2, "j"
    End If
End Sub

Public Sub CalculateSW_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 3, "k"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 3, "g"
    End If
End Sub

Public Sub CalculateSE_Waterfall(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile As Boolean
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile = True
    
    ' Actually place the subtile
    If tmpTile Then
        ' Extended
        placeAutotile layerNum, X, Y, 4, "h"
    Else
        ' Edge
        placeAutotile layerNum, X, Y, 4, "l"
    End If
End Sub

' Cliff autotiling
Public Sub CalculateNW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North West
    If checkTileMatch(layerNum, X, Y, X - 1, Y - 1) Then tmpTile(1) = True
    
    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(2) = True
    
    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(2) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(2) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(2) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 1, "e"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 1, "i"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 1, "m"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 1, "q"
    End Select
End Sub

Public Sub CalculateNE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' North
    If checkTileMatch(layerNum, X, Y, X, Y - 1) Then tmpTile(1) = True
    
    ' North East
    If checkTileMatch(layerNum, X, Y, X + 1, Y - 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 2, "j"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 2, "f"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 2, "r"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 2, "n"
    End Select
End Sub

Public Sub CalculateSW_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' West
    If checkTileMatch(layerNum, X, Y, X - 1, Y) Then tmpTile(1) = True
    
    ' South West
    If checkTileMatch(layerNum, X, Y, X - 1, Y + 1) Then tmpTile(2) = True
    
    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(3) = True
    
    ' Calculate Situation - Horizontal
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 3, "o"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 3, "s"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 3, "g"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 3, "k"
    End Select
End Sub

Public Sub CalculateSE_Cliff(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim tmpTile(1 To 3) As Boolean
    Dim situation As Byte

    ' South
    If checkTileMatch(layerNum, X, Y, X, Y + 1) Then tmpTile(1) = True
    
    ' South East
    If checkTileMatch(layerNum, X, Y, X + 1, Y + 1) Then tmpTile(2) = True
    
    ' East
    If checkTileMatch(layerNum, X, Y, X + 1, Y) Then tmpTile(3) = True
    
    ' Calculate Situation -  Horizontal
    If Not tmpTile(1) And tmpTile(3) Then situation = AUTO_HORIZONTAL
    ' Vertical
    If tmpTile(1) And Not tmpTile(3) Then situation = AUTO_VERTICAL
    ' Fill
    If tmpTile(1) And tmpTile(2) And tmpTile(3) Then situation = AUTO_FILL
    ' Inner
    If Not tmpTile(1) And Not tmpTile(3) Then situation = AUTO_INNER
    
    ' Actually place the subtile
    Select Case situation
        Case AUTO_INNER
            placeAutotile layerNum, X, Y, 4, "t"
        Case AUTO_HORIZONTAL
            placeAutotile layerNum, X, Y, 4, "p"
        Case AUTO_VERTICAL
            placeAutotile layerNum, X, Y, 4, "l"
        Case AUTO_FILL
            placeAutotile layerNum, X, Y, 4, "h"
    End Select
End Sub

Public Function checkTileMatch(ByVal layerNum As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    ' we'll exit out early if true
    checkTileMatch = True
    
    ' If it's off the map then set it as autotile and exit out early
    If X2 < 0 Or X2 > Map.MaxX Or Y2 < 0 Or Y2 > Map.MaxY Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' Fakes ALWAYS return true
    If Map.Tile(X2, Y2).Autotile(layerNum) = AUTOTILE_FAKE Then
        checkTileMatch = True
        Exit Function
    End If
    
    ' Check neighbour is an autotile
    If Map.Tile(X2, Y2).Autotile(layerNum) = 0 Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' Check we're a matching
    If Map.Tile(X1, Y1).Layer(layerNum).Tileset <> Map.Tile(X2, Y2).Layer(layerNum).Tileset Then
        checkTileMatch = False
        Exit Function
    End If
    
    ' Check tiles match
    If Map.Tile(X1, Y1).Layer(layerNum).X <> Map.Tile(X2, Y2).Layer(layerNum).X Then
        checkTileMatch = False
        Exit Function
    End If
        
    If Map.Tile(X1, Y1).Layer(layerNum).Y <> Map.Tile(X2, Y2).Layer(layerNum).Y Then
        checkTileMatch = False
        Exit Function
    End If
End Function

'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'   All of this code is for auto tiles and the math behind generating them.
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
Public Sub placeAutotile(ByVal layerNum As Long, ByVal X As Long, ByVal Y As Long, ByVal tileQuarter As Byte, ByVal autoTileLetter As String)
    With Autotile(X, Y).Layer(layerNum).QuarterTile(tileQuarter)
        Select Case autoTileLetter
            Case "a"
                .X = autoInner(1).X
                .Y = autoInner(1).Y
            Case "b"
                .X = autoInner(2).X
                .Y = autoInner(2).Y
            Case "c"
                .X = autoInner(3).X
                .Y = autoInner(3).Y
            Case "d"
                .X = autoInner(4).X
                .Y = autoInner(4).Y
            Case "e"
                .X = autoNW(1).X
                .Y = autoNW(1).Y
            Case "f"
                .X = autoNW(2).X
                .Y = autoNW(2).Y
            Case "g"
                .X = autoNW(3).X
                .Y = autoNW(3).Y
            Case "h"
                .X = autoNW(4).X
                .Y = autoNW(4).Y
            Case "i"
                .X = autoNE(1).X
                .Y = autoNE(1).Y
            Case "j"
                .X = autoNE(2).X
                .Y = autoNE(2).Y
            Case "k"
                .X = autoNE(3).X
                .Y = autoNE(3).Y
            Case "l"
                .X = autoNE(4).X
                .Y = autoNE(4).Y
            Case "m"
                .X = autoSW(1).X
                .Y = autoSW(1).Y
            Case "n"
                .X = autoSW(2).X
                .Y = autoSW(2).Y
            Case "o"
                .X = autoSW(3).X
                .Y = autoSW(3).Y
            Case "p"
                .X = autoSW(4).X
                .Y = autoSW(4).Y
            Case "q"
                .X = autoSE(1).X
                .Y = autoSE(1).Y
            Case "r"
                .X = autoSE(2).X
                .Y = autoSE(2).Y
            Case "s"
                .X = autoSE(3).X
                .Y = autoSE(3).Y
            Case "t"
                .X = autoSE(4).X
                .Y = autoSE(4).Y
        End Select
    End With
End Sub

Public Sub DrawFog()
    Dim fogNum As Long, Color As Long, X As Long, Y As Long, RenderState As Long

    fogNum = CurrentFog
    If fogNum <= 0 Or fogNum > NumFogs Then Exit Sub
    Color = D3DColorRGBA(255, 255, 255, CurrentFogOpacity)

    RenderState = 0
    
    ' Render state
    Select Case RenderState
        Case 1 ' Additive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For X = 0 To ((Map.MaxX * 32) / 256) + 1
        For Y = 0 To ((Map.MaxY * 32) / 256) + 1
            RenderTexture Tex_Fog(fogNum), ConvertMapX((X * 256) + fogOffsetX), ConvertMapY((Y * 256) + fogOffsetY), 0, 0, 256, 256, 256, 256, Color
        Next
    Next
    
    ' Reset render state
    If RenderState > 0 Then
        Direct3D_Device.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        Direct3D_Device.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        Direct3D_Device.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawTint()
    Dim Color As Long
    
    Color = D3DColorRGBA(CurrentTintR, CurrentTintG, CurrentTintB, CurrentTintA)
    
    RenderTexture Tex_White, 0, 0, 0, 0, frmMain.picScreen.ScaleWidth, frmMain.picScreen.ScaleHeight, 32, 32, Color
End Sub

Public Sub DrawWeather()
    Dim Color As Long, I As Long, SpriteLeft As Long
    
    For I = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(I).InUse Then
            If WeatherParticle(I).Type = WEATHER_TYPE_STORM Then
                SpriteLeft = 0
            Else
                SpriteLeft = WeatherParticle(I).Type - 1
            End If
            RenderTexture Tex_Weather, ConvertMapX(WeatherParticle(I).X), ConvertMapY(WeatherParticle(I).Y), SpriteLeft * 32, 0, 32, 32, 32, 32, -1
        End If
    Next
End Sub

Public Sub EditorMapProperties_DrawPanorama()
    Dim Height As Long, srcRect As D3DRECT, destRECT As D3DRECT
    Dim Width As Long
    Dim Panorama As Long
    Dim sRECT As RECT
    Dim dRect As RECT
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Find Panorama number
    Panorama = frmEditor_MapProperties.scrlPanorama.Value
    
    ' Exit out if doesn't exist
    If Panorama < 1 Or Panorama > NumPanoramas Then
        frmEditor_MapProperties.picPanorama.Cls
        Exit Sub
    End If
    
    Direct3D_Device.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorRGBA(0, 0, 0, 0), 1#, 0
    Direct3D_Device.BeginScene
    
    Height = Tex_Panorama(Panorama).Height
    Width = Tex_Panorama(Panorama).Width
    
    sRECT.Left = 0
    sRECT.Top = 0
    sRECT.Right = sRECT.Left + Width
    sRECT.Bottom = sRECT.Top + Height
    
    dRect.Top = 0
    dRect.Bottom = Height
    dRect.Left = 0
    dRect.Right = Width
    
    RenderTextureByRects Tex_Panorama(Panorama), sRECT, dRect
                    
    With destRECT
        .X1 = 0
        .X2 = frmEditor_MapProperties.picPanorama.Width
        .Y1 = 0
        .Y2 = frmEditor_MapProperties.picPanorama.Height
    End With
                
    Direct3D_Device.EndScene
    Direct3D_Device.Present destRECT, destRECT, frmEditor_MapProperties.picPanorama.hWnd, ByVal (0)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "EditorMapProperties_DrawPanorama", "modRendering", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
