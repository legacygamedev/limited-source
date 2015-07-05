Attribute VB_Name = "modDirectX8"
Option Explicit

' Texture wrapper
Public Tex_Anim() As Long
Public Tex_Char() As Long
Public Tex_Face() As Long
Public Tex_GUI() As Long
Public Tex_Item() As Long
Public Tex_Paperdoll() As Long
Public Tex_Resource() As Long
Public Tex_Spellicon() As Long
Public Tex_Tileset() As Long
Public Tex_Buttons() As Long
Public Tex_Buttons_h() As Long
Public Tex_Buttons_c() As Long
Public Tex_Fog() As Long
Public Tex_Bars As Long
Public Tex_Blood As Long
Public Tex_Direction As Long
Public Tex_Misc As Long
Public Tex_Target As Long
Public Tex_Shadow As Long
Public Tex_Fader As Long

' Surfaces
Public Tex_Surface() As SurfaceRec
Public Count_Surface As Long

' Texture count
Public Count_Anim As Long
Public Count_Char As Long
Public Count_Face As Long
Public Count_GUI As Long
Public Count_Item As Long
Public Count_Paperdoll As Long
Public Count_Resource As Long
Public Count_Spellicon As Long
Public Count_Tileset As Long
Public Count_Fog As Long

' Texture paths
Public Const Path_Anim As String = "\data files\graphics\animations\"
Public Const Path_Char As String = "\data files\graphics\characters\"
Public Const Path_Face As String = "\data files\graphics\faces\"
Public Const Path_GUI As String = "\data files\graphics\gui\"
Public Const Path_Item As String = "\data files\graphics\items\"
Public Const Path_Paperdoll As String = "\data files\graphics\paperdolls\"
Public Const Path_Resource As String = "\data files\graphics\resources\"
Public Const Path_Spellicon As String = "\data files\graphics\spellicons\"
Public Const Path_Tileset As String = "\data files\graphics\tilesets\"
Public Const Path_Font As String = "\data files\graphics\fonts\"
Public Const Path_Graphics As String = "\data files\graphics\"
Public Const Path_Buttons As String = "\data files\graphics\gui\buttons\"
Public Const Path_Surface As String = "\data files\graphics\surfaces\"
Public Const Path_Fog As String = "\data files\graphics\fog\"

' Variables
Public dX As DirectX8
Public D3D8 As Direct3D8
Public Direct3DX8 As D3DX8

Public D3DDevice8 As Direct3DDevice8
Public DispMode As D3DDISPLAYMODE
Public D3DWindow As D3DPRESENT_PARAMETERS
Public BackBuffer As Direct3DSurface8

Public Type TLVERTEX
    x As Single
    y As Single
    z As Single
    RHW As Single
    Color As Long
    tu As Single
    tv As Single
End Type

Public Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE
Public Const FVF_Size As Long = 28

Public D3DT_TEXTURE() As TextureRec
Private Const TEXTURE_NULL As Long = 0

Private Type TextureRec
    Texture As Direct3DTexture8
    Width As Long
    height As Long
    path As String
    UnloadTimer As Long
    loaded As Boolean
End Type

Private Type SurfaceRec
    Surface As Direct3DSurface8
    Width As Long
    height As Long
    path As String
    UnloadTimer As Long
    loaded As Boolean
End Type

Private Type D3DXIMAGE_INFO_A
    Width As Long
    height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
End Type

Public Type GeomRec
    top As Long
    left As Long
    height As Long
    Width As Long
End Type

Private mTextureNum As Long
Public CurrentTexture As Long

Public ScreenWidth As Long
Public ScreenHeight As Long

Public Const DegreeToRadian As Single = 0.0174532919296
Public Const RadianToDegree As Single = 57.2958300962816

Public Sub EngineInit()
    
    Set dX = New DirectX8
    Set D3D8 = dX.Direct3DCreate()
    Set Direct3DX8 = New D3DX8
    
    ' hardware acceleration
    Select Case Options.render
        Case 1 ' software
            If Not EngineInitD3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                Call MsgBox("Could not init D3DDevice8. Exiting...")
                Call EngineUnloadDirectX
                DestroyGame
                End
            End If
        Case 2 ' mixed
            If Not EngineInitD3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                Call MsgBox("Could not init D3DDevice8. Exiting...")
                Call EngineUnloadDirectX
                DestroyGame
                End
            End If
        Case 3 ' hardware
            If Not EngineInitD3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
                Call MsgBox("Could not init D3DDevice8. Exiting...")
                Call EngineUnloadDirectX
                DestroyGame
                End
            End If
        Case 4 ' pure device
            If Not EngineInitD3DDevice(D3DCREATE_PUREDEVICE Or D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
                Call MsgBox("Could not init D3DDevice8. Exiting...")
                Call EngineUnloadDirectX
                DestroyGame
                End
            End If
        Case Else ' optimal
            If Not EngineInitD3DDevice(D3DCREATE_PUREDEVICE Or D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
                If Not EngineInitD3DDevice(D3DCREATE_HARDWARE_VERTEXPROCESSING) Then
                    If Not EngineInitD3DDevice(D3DCREATE_MIXED_VERTEXPROCESSING) Then
                        If Not EngineInitD3DDevice(D3DCREATE_SOFTWARE_VERTEXPROCESSING) Then
                            Call MsgBox("Could not init D3DDevice8. Exiting...")
                            Call EngineUnloadDirectX
                            DestroyGame
                            End
                        End If
                    End If
                End If
            End If
    End Select

    Call EngineCacheTextures
    Call EngineInitRenderStates
    Call EngineInitFontTextures
    Call EngineInitFontSettings
    Call UpdateChatArray
End Sub

Public Function EngineInitRenderStates()
    With D3DDevice8
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
End Function

Public Function EngineInitD3DDevice(D3DCREATEFLAGS As CONST_D3DCREATEFLAGS) As Boolean
On Error GoTo ERRORMSG
    
    D3D8.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    ScreenWidth = 800
    ScreenHeight = 600
    
    DispMode.Format = D3DFMT_X8R8G8B8
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    DispMode.Width = ScreenWidth
    DispMode.height = ScreenHeight
    D3DWindow.BackBufferCount = 1
    D3DWindow.BackBufferFormat = DispMode.Format
    D3DWindow.BackBufferWidth = ScreenWidth
    D3DWindow.BackBufferHeight = ScreenHeight
    D3DWindow.hDeviceWindow = frmMain.hWnd
    D3DWindow.Windowed = True

    If Not D3DDevice8 Is Nothing Then Set D3DDevice8 = Nothing
    Set D3DDevice8 = D3D8.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, D3DCREATEFLAGS, D3DWindow)
    
    ' set the backbuffer
    Set BackBuffer = D3DDevice8.GetBackBuffer(0, D3DBACKBUFFER_TYPE_MONO)
    
    EngineInitD3DDevice = True
    Exit Function
    
ERRORMSG:
    Set D3DDevice8 = Nothing
    EngineInitD3DDevice = False
End Function

Public Sub EngineUnloadDirectX()
Dim i As Long

    If Not D3DDevice8 Is Nothing Then Set D3DDevice8 = Nothing
    If Not D3D8 Is Nothing Then Set D3D8 = Nothing

    For i = 1 To mTextureNum
        Set D3DT_TEXTURE(i).Texture = Nothing
    Next

    If Not dX Is Nothing Then Set dX = Nothing
End Sub

Public Sub EngineCacheTextures()
Dim i As Long

    ' Animation Textures
    Count_Anim = 1
    Do While FileExist(App.path & Path_GUI & Count_Anim & ".png")
        ReDim Preserve Tex_Anim(0 To Count_Anim)
        Tex_Anim(Count_Anim) = SetTexturePath(App.path & Path_Anim & Count_Anim & ".png")
        Count_Anim = Count_Anim + 1
    Loop
    Count_Anim = Count_Anim - 1
    
    ' Character Textures
    Count_Char = 1
    Do While FileExist(App.path & Path_Char & Count_Char & ".png")
        ReDim Preserve Tex_Char(0 To Count_Char)
        Tex_Char(Count_Char) = SetTexturePath(App.path & Path_Char & Count_Char & ".png")
        Count_Char = Count_Char + 1
    Loop
    Count_Char = Count_Char - 1
    
    ' Face Textures
    Count_Face = 1
    Do While FileExist(App.path & Path_Face & Count_Face & ".png")
        ReDim Preserve Tex_Face(0 To Count_Face)
        Tex_Face(Count_Face) = SetTexturePath(App.path & Path_Face & Count_Face & ".png")
        Count_Face = Count_Face + 1
    Loop
    Count_Face = Count_Face - 1
    
    ' GUI Textures
    Count_GUI = 1
    Do While FileExist(App.path & Path_GUI & Count_GUI & ".png")
        ReDim Preserve Tex_GUI(0 To Count_GUI)
        Tex_GUI(Count_GUI) = SetTexturePath(App.path & Path_GUI & Count_GUI & ".png")
        Count_GUI = Count_GUI + 1
    Loop
    Count_GUI = Count_GUI - 1
    
    ' Item Textures
    Count_Item = 1
    Do While FileExist(App.path & Path_Item & Count_Item & ".png")
        ReDim Preserve Tex_Item(0 To Count_Item)
        Tex_Item(Count_Item) = SetTexturePath(App.path & Path_Item & Count_Item & ".png")
        Count_Item = Count_Item + 1
    Loop
    Count_Item = Count_Item - 1

    ' Paperdoll Textures
    Count_Paperdoll = 1
    Do While FileExist(App.path & Path_Paperdoll & Count_Paperdoll & ".png")
        ReDim Preserve Tex_Paperdoll(0 To Count_Paperdoll)
        Tex_Paperdoll(Count_Paperdoll) = SetTexturePath(App.path & Path_Paperdoll & Count_Paperdoll & ".png")
        Count_Paperdoll = Count_Paperdoll + 1
    Loop
    Count_Paperdoll = Count_Paperdoll - 1

    ' Resource Textures
    Count_Resource = 1
    Do While FileExist(App.path & Path_Resource & Count_Resource & ".png")
        ReDim Preserve Tex_Resource(0 To Count_Resource)
        Tex_Resource(Count_Resource) = SetTexturePath(App.path & Path_Resource & Count_Resource & ".png")
        Count_Resource = Count_Resource + 1
    Loop
    Count_Resource = Count_Resource - 1

    ' SpellIcon Textures
    Count_Spellicon = 1
    Do While FileExist(App.path & Path_Spellicon & Count_Spellicon & ".png")
        ReDim Preserve Tex_Spellicon(0 To Count_Spellicon)
        Tex_Spellicon(Count_Spellicon) = SetTexturePath(App.path & Path_Spellicon & Count_Spellicon & ".png")
        Count_Spellicon = Count_Spellicon + 1
    Loop
    Count_Spellicon = Count_Spellicon - 1

    ' Tileset Textures
    Count_Tileset = 1
    Do While FileExist(App.path & Path_Tileset & Count_Tileset & ".png")
        ReDim Preserve Tex_Tileset(0 To Count_Tileset)
        Tex_Tileset(Count_Tileset) = SetTexturePath(App.path & Path_Tileset & Count_Tileset & ".png")
        Count_Tileset = Count_Tileset + 1
    Loop
    Count_Tileset = Count_Tileset - 1

    ' Buttons
    ReDim Tex_Buttons(1 To MAX_BUTTONS)
    ReDim Tex_Buttons_h(1 To MAX_BUTTONS)
    ReDim Tex_Buttons_c(1 To MAX_BUTTONS)
    For i = 1 To MAX_BUTTONS
        Tex_Buttons(i) = SetTexturePath(App.path & Path_Buttons & i & ".png")
        Tex_Buttons_h(i) = SetTexturePath(App.path & Path_Buttons & i & "_h.png")
        Tex_Buttons_c(i) = SetTexturePath(App.path & Path_Buttons & i & "_c.png")
    Next
    
    ' Fog Textures
    Count_Fog = 1
    Do While FileExist(App.path & Path_Fog & Count_Fog & ".png")
        ReDim Preserve Tex_Fog(0 To Count_Fog)
        Tex_Fog(Count_Fog) = SetTexturePath(App.path & Path_Fog & Count_Fog & ".png")
        Count_Fog = Count_Fog + 1
    Loop
    Count_Fog = Count_Fog - 1
    
    ' Surfaces
    Count_Surface = 1
    Do While FileExist(App.path & Path_Surface & Count_Surface & ".png")
        ReDim Preserve Tex_Surface(0 To Count_Surface)
        Tex_Surface(Count_Surface).path = App.path & Path_Surface & Count_Surface & ".png"
        Count_Surface = Count_Surface + 1
    Loop
    Count_Surface = Count_Surface - 1
    
    ' Single Textures
    Tex_Bars = SetTexturePath(App.path & Path_Graphics & "bars.png")
    Tex_Blood = SetTexturePath(App.path & Path_Graphics & "blood.png")
    Tex_Direction = SetTexturePath(App.path & Path_Graphics & "direction.png")
    Tex_Misc = SetTexturePath(App.path & Path_Graphics & "misc.png")
    Tex_Target = SetTexturePath(App.path & Path_Graphics & "target.png")
    Tex_Fader = SetTexturePath(App.path & Path_Graphics & "fader.png")
    Tex_Shadow = SetTexturePath(App.path & Path_Graphics & "shadow.png")
End Sub

Public Function SetTexturePath(ByVal path As String) As Long
    mTextureNum = mTextureNum + 1
    ReDim Preserve D3DT_TEXTURE(0 To mTextureNum) As TextureRec
    D3DT_TEXTURE(mTextureNum).path = path
    SetTexturePath = mTextureNum
    D3DT_TEXTURE(mTextureNum).loaded = False
End Function

Public Sub LoadTexture(ByVal TextureNum As Long)
Dim Tex_Info As D3DXIMAGE_INFO_A
Dim path As String

    path = D3DT_TEXTURE(TextureNum).path
    
    Select Case D3DT_TEXTURE(TextureNum).Width
        Case 0
            Set D3DT_TEXTURE(TextureNum).Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, path, D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, RGB(255, 0, 255), Tex_Info, ByVal 0)
            D3DT_TEXTURE(TextureNum).height = Tex_Info.height
            D3DT_TEXTURE(TextureNum).Width = Tex_Info.Width
        Case Is > 0
            Set D3DT_TEXTURE(TextureNum).Texture = Direct3DX8.CreateTextureFromFileEx(D3DDevice8, path, D3DT_TEXTURE(TextureNum).Width, D3DT_TEXTURE(TextureNum).height, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, RGB(255, 0, 255), ByVal 0, ByVal 0)
    End Select
    
    D3DT_TEXTURE(TextureNum).UnloadTimer = GetTickCount
    D3DT_TEXTURE(TextureNum).loaded = True
End Sub

Public Sub LoadSurface(ByVal Width As Long, ByVal height As Long, ByVal SurfaceNum As Long)
Dim path As String
    
    path = Tex_Surface(SurfaceNum).path
    
    Set Tex_Surface(SurfaceNum).Surface = D3DDevice8.CreateImageSurface(Width, height, DispMode.Format)
    Direct3DX8.LoadSurfaceFromFile Tex_Surface(SurfaceNum).Surface, ByVal 0, ByVal 0, path, ByVal 0, D3DX_DEFAULT, 0, ByVal 0
    
    Tex_Surface(SurfaceNum).Width = Width
    Tex_Surface(SurfaceNum).height = height
    Tex_Surface(SurfaceNum).loaded = True
    Tex_Surface(SurfaceNum).UnloadTimer = GetTickCount
End Sub

Public Sub UnloadTextures()
Dim Count As Long
Dim i As Long

    Count = UBound(D3DT_TEXTURE)
    If Count <= 0 Then Exit Sub
    
    For i = 1 To Count
        With D3DT_TEXTURE(i)
            If .UnloadTimer > GetTickCount + 150000 Then
                Set .Texture = Nothing
                Call ZeroMemory(ByVal VarPtr(D3DT_TEXTURE(i)), LenB(D3DT_TEXTURE(i)))
                .UnloadTimer = 0
                .loaded = False
            End If
        End With
    Next
End Sub

Public Sub SetTexture(ByVal Texture As Long)
    If Texture <> CurrentTexture Then
    
        If Texture > UBound(D3DT_TEXTURE) Then Texture = UBound(D3DT_TEXTURE)
        If Texture < 0 Then Texture = 0
        
        If Not Texture = TEXTURE_NULL Then
            If Not D3DT_TEXTURE(Texture).loaded Then
                Call LoadTexture(Texture)
            End If
        End If
        
        Call D3DDevice8.SetTexture(0, D3DT_TEXTURE(Texture).Texture)
        CurrentTexture = Texture
    End If
End Sub

Public Sub EngineRenderSurface(ByVal Surface As Long, ByVal dX As Long, ByVal dY As Long, ByVal sX As Long, ByVal sY As Long, ByVal sW As Long, ByVal sH As Long)
Dim sRECT As DxVBLibA.RECT
Dim dPOS As DxVBLibA.Point

    ' make sure surface exists
    If Surface < 0 Or Surface > Count_Surface Then Exit Sub
    
    ' load it if we need to
    If Not Tex_Surface(Surface).loaded Then
        Call LoadSurface(sW, sH, Surface)
    End If

    With sRECT
        .top = sY
        .left = sX
        .Right = .left + sW
        .bottom = .top + sH
    End With
    
    With dPOS
        .x = dX
        .y = dY
    End With

    D3DDevice8.CopyRects Tex_Surface(Surface).Surface, sRECT, 1, BackBuffer, dPOS
End Sub

Public Sub RenderTexture(ByVal Texture As Long, ByVal dX As Long, ByVal dY As Long, ByVal sX As Long, ByVal sY As Long, ByVal dW As Long, ByVal dH As Long, ByVal sW As Long, ByVal sH As Long, Optional ByVal colour As Long = -1)
Dim Box(0 To 3) As TLVERTEX, x As Long, textureWidth As Long, textureHeight As Long

    ' set the texture
    Call SetTexture(Texture)
    
    ' set the texture size
    textureWidth = D3DT_TEXTURE(Texture).Width
    textureHeight = D3DT_TEXTURE(Texture).height
    
    ' exit out if we need to
    If Texture <= 0 Or textureWidth <= 0 Or textureHeight <= 0 Then Exit Sub
    
    For x = 0 To 3
        Box(x).RHW = 1
        Box(x).Color = colour
    Next

    Box(0).x = dX
    Box(0).y = dY
    Box(0).tu = (sX / textureWidth)
    Box(0).tv = (sY / textureHeight)
    Box(1).x = dX + dW
    Box(1).tu = (sX + sW + 1) / textureWidth
    Box(2).x = Box(0).x
    Box(3).x = Box(1).x

    Box(2).y = dY + dH
    Box(2).tv = (sY + sH + 1) / textureHeight

    Box(1).y = Box(0).y
    Box(1).tv = Box(0).tv
    Box(2).tu = Box(0).tu
    Box(3).y = Box(2).y
    Box(3).tu = Box(1).tu
    Box(3).tv = Box(2).tv
    
    Call D3DDevice8.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), FVF_Size)
    D3DT_TEXTURE(Texture).UnloadTimer = GetTickCount
End Sub

Public Sub EngineDrawSquare(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, Optional ByVal colour As Long = 0, Optional lineWidth As Byte = 1)
    ' Top line
    'EngineRenderRectangle 0, x, y, 0, 0, w, lineWidth, , , , , colour, colour, colour, colour
    RenderTexture 0, x, y, 0, 0, w, lineWidth, 0, 0, colour
    ' Left line
    'EngineRenderRectangle 0, x, y, 0, 0, lineWidth, h, , , , , colour, colour, colour, colour
    RenderTexture 0, x, y, 0, 0, lineWidth, h, 0, 0, colour
    ' Bottom line
    'EngineRenderRectangle 0, x, y + h - lineWidth, 0, 0, w, lineWidth, , , , , colour, colour, colour, colour
    RenderTexture 0, x, y + h - lineWidth, 0, 0, w, lineWidth, 0, 0, colour
    ' Right line
    'EngineRenderRectangle 0, x + w - lineWidth, y, 0, 0, lineWidth, h, , , , , colour, colour, colour, colour
    RenderTexture 0, x + w - lineWidth, y, 0, 0, lineWidth, h, 0, 0, colour
End Sub

' GDI rendering
Public Sub GDIRenderAnimation()
Dim i As Long, Animationnum As Long, ShouldRender As Boolean, Width As Long, height As Long, looptime As Long, FrameCount As Long
Dim sX As Long, sY As Long, sRECT As RECT

    sRECT.top = 0
    sRECT.bottom = 192
    sRECT.left = 0
    sRECT.Right = 192

    For i = 0 To 1
        Animationnum = frmEditor_Animation.scrlSprite(i).value
        
        If Animationnum <= 0 Or Animationnum > Count_Anim Then
            ' don't render lol
        Else
            looptime = frmEditor_Animation.scrlLoopTime(i)
            FrameCount = frmEditor_Animation.scrlFrameCount(i)
            
            ShouldRender = False
            
            ' check if we need to render new frame
            If AnimEditorTimer(i) + looptime <= GetTickCount Then
                ' check if out of range
                If AnimEditorFrame(i) >= FrameCount Then
                    AnimEditorFrame(i) = 1
                Else
                    AnimEditorFrame(i) = AnimEditorFrame(i) + 1
                End If
                AnimEditorTimer(i) = GetTickCount
                ShouldRender = True
            End If
        
            If ShouldRender Then
                If frmEditor_Animation.scrlFrameCount(i).value > 0 Then
                    ' total width divided by frame count
                    Width = 192
                    height = 192

                    sY = (height * ((AnimEditorFrame(i) - 1) \ AnimColumns))
                    sX = (Width * (((AnimEditorFrame(i) - 1) Mod AnimColumns)))

                    ' Start Rendering
                    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
                    Call D3DDevice8.BeginScene
                    
                    'EngineRenderRectangle Tex_Anim(Animationnum), 0, 0, sX, sY, width, height, width, height
                    RenderTexture Tex_Anim(Animationnum), 0, 0, sX, sY, Width, height, Width, height
                    
                    ' Finish Rendering
                    Call D3DDevice8.EndScene
                    Call D3DDevice8.Present(sRECT, ByVal 0, frmEditor_Animation.picSprite(i).hWnd, ByVal 0)
                End If
            End If
        End If
    Next
End Sub

Public Sub GDIRenderChar(ByRef picBox As PictureBox, ByVal sprite As Long)
Dim height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Char Then Exit Sub
    
    height = 32
    Width = 32
    
    sRECT.top = 0
    sRECT.bottom = sRECT.top + height
    sRECT.left = 0
    sRECT.Right = sRECT.left + Width

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    RenderTexture Tex_Char(sprite), 0, 0, 0, 0, Width, height, Width, height
     
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderFace(ByRef picBox As PictureBox, ByVal sprite As Long)
Dim height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Face Then Exit Sub
    
    height = D3DT_TEXTURE(Tex_Face(sprite)).height
    Width = D3DT_TEXTURE(Tex_Face(sprite)).Width
    
    If height = 0 Or Width = 0 Then
        height = 1
        Width = 1
    End If
    
    sRECT.top = 0
    sRECT.bottom = sRECT.top + height
    sRECT.left = 0
    sRECT.Right = sRECT.left + Width
    

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Face(sprite), 0, 0, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Face(sprite), 0, 0, 0, 0, Width, height, Width, height
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderTileset()
Dim height As Long, Width As Long, Tileset As Byte, sRECT As RECT

    ' find tileset number
    Tileset = frmEditor_Map.scrlTileSet.value
    
    ' exit out if doesn't exist
    If Tileset <= 0 Or Tileset > Count_Tileset Then Exit Sub
    
    height = D3DT_TEXTURE(Tex_Tileset(Tileset)).height
    Width = D3DT_TEXTURE(Tex_Tileset(Tileset)).Width
    
    If height = 0 Or Width = 0 Then
        height = 1
        Width = 1
    End If
    
    frmEditor_Map.picBackSelect.Width = Width
    frmEditor_Map.picBackSelect.height = height
    
    sRECT.top = 0
    sRECT.bottom = height
    sRECT.left = 0
    sRECT.Right = Width
    
    ' change selected shape for autotiles
    If frmEditor_Map.scrlAutotile.value > 0 Then
        Select Case frmEditor_Map.scrlAutotile.value
            Case 1 ' autotile
                shpSelectedWidth = 64
                shpSelectedHeight = 96
            Case 2 ' fake autotile
                shpSelectedWidth = 32
                shpSelectedHeight = 32
            Case 3 ' animated
                shpSelectedWidth = 192
                shpSelectedHeight = 96
            Case 4 ' cliff
                shpSelectedWidth = 64
                shpSelectedHeight = 64
            Case 5 ' waterfall
                shpSelectedWidth = 64
                shpSelectedHeight = 96
        End Select
    End If

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, dx8Colour(White, 255), 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Tileset(Tileset), 0, 0, 0, 0, width, height, width, height, width, height
    If Tex_Tileset(Tileset) <= 0 Then Exit Sub
    RenderTexture Tex_Tileset(Tileset), 0, 0, 0, 0, Width, height, Width, height
    
    ' draw selection boxes
    EngineDrawSquare shpSelectedLeft, shpSelectedTop, shpSelectedWidth, shpSelectedHeight, dx8Colour(Red, 255), 2
    
    ' draw selection boxes
    EngineDrawSquare shpLocLeft, shpLocTop, 32, 32, dx8Colour(Blue, 255), 2
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, frmEditor_Map.picBackSelect.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderItem(ByRef picBox As PictureBox, ByVal sprite As Long)
Dim height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Item Then Exit Sub
    
    height = D3DT_TEXTURE(Tex_Item(sprite)).height
    Width = D3DT_TEXTURE(Tex_Item(sprite)).Width
    
    sRECT.top = 0
    sRECT.bottom = 32
    sRECT.left = 0
    sRECT.Right = 32

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(sprite), 0, 0, 0, 0, 32, 32, 32, 32
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

Public Sub GDIRenderSpell(ByRef picBox As PictureBox, ByVal sprite As Long)
Dim height As Long, Width As Long, sRECT As RECT

    ' exit out if doesn't exist
    If sprite <= 0 Or sprite > Count_Spellicon Then Exit Sub
    
    height = D3DT_TEXTURE(Tex_Spellicon(sprite)).height
    Width = D3DT_TEXTURE(Tex_Spellicon(sprite)).Width
    
    If height = 0 Or Width = 0 Then
        height = 1
        Width = 1
    End If
    
    sRECT.top = 0
    sRECT.bottom = height
    sRECT.left = 0
    sRECT.Right = Width

    ' Start Rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    'EngineRenderRectangle Tex_Spellicon(sprite), 0, 0, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Spellicon(sprite), 0, 0, 0, 0, 32, 32, 32, 32
    
    ' Finish Rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(sRECT, ByVal 0, picBox.hWnd, ByVal 0)
End Sub

' Directional blocking
Public Sub DrawDirection(ByVal x As Long, ByVal y As Long)
Dim i As Long, top As Long, left As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' render grid
    top = 24
    left = 0
    'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Direction, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), left, top, 32, 32, 32, 32
    
    ' render dir blobs
    For i = 1 To 4
        left = (i - 1) * 8
        ' find out whether render blocked or not
        If Not isDirBlocked(Map.Tile(x, y).DirBlock, CByte(i)) Then
            top = 8
        Else
            top = 16
        End If
        'render!
        'EngineRenderRectangle Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8, 8, 8
        RenderTexture Tex_Direction, ConvertMapX(x * PIC_X) + DirArrowX(i), ConvertMapY(y * PIC_Y) + DirArrowY(i), left, top, 8, 8, 8, 8
    Next
    
    ' Error handler
    Exit Sub
errorhandler:
    HandleError "BltDirection", "modDirectX8", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Public Sub DrawFog()
Dim fogNum As Long, colour As Long, x As Long, y As Long, renderState As Long
    
    fogNum = 3
    If fogNum <= 0 Or fogNum > Count_Fog Then Exit Sub
    colour = D3DColorARGB(64, 255, 255, 255)
    renderState = 0
    
    Exit Sub
    
    ' render state
    Select Case renderState
        Case 1 ' Additive
            D3DDevice8.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
            D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
        Case 2 ' Subtractive
            D3DDevice8.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_SUBTRACT
            D3DDevice8.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ZERO
            D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR
    End Select
    
    For x = 0 To 4
        For y = 0 To 3
            'RenderTexture Tex_Fog(fogNum), (x * 256) + fogOffsetX, (y * 256) + fogOffsetY, 0, 0, 256, 256, 256, 256, colour
            RenderTexture Tex_Fog(fogNum), (x * 256), (y * 256), 0, 0, 256, 256, 256, 256, colour
        Next
    Next
    
    ' reset render state
    If renderState > 0 Then
        D3DDevice8.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        D3DDevice8.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        D3DDevice8.SetTextureStageState 0, D3DTSS_COLOROP, D3DTOP_MODULATE
    End If
End Sub

Public Sub DrawAutoTile(ByVal layerNum As Long, ByVal destX As Long, ByVal destY As Long, ByVal quarterNum As Long, ByVal x As Long, ByVal y As Long)
Dim yOffset As Long, xOffset As Long

    ' calculate the offset
    Select Case Map.Tile(x, y).Autotile(layerNum)
        Case AUTOTILE_WATERFALL
            yOffset = (waterfallFrame - 1) * 32
        Case AUTOTILE_ANIM
            xOffset = autoTileFrame * 64
        Case AUTOTILE_CLIFF
            yOffset = -32
    End Select
    
    ' Draw the quarter
    'EngineRenderRectangle Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16, 16, 16
    RenderTexture Tex_Tileset(Map.Tile(x, y).Layer(layerNum).Tileset), destX, destY, Autotile(x, y).Layer(layerNum).srcX(quarterNum) + xOffset, Autotile(x, y).Layer(layerNum).srcY(quarterNum) + yOffset, 16, 16, 16, 16
End Sub

' Rendering Procedures
Public Sub DrawMapTile(ByVal x As Long, ByVal y As Long)
Dim i As Long
    
    With Map.Tile(x, y)
        ' draw the map
        For i = MapLayer.Ground To MapLayer.Mask2
            ' skip tile if tileset isn't set
            If Autotile(x, y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                'EngineRenderRectangle Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
            ElseIf Autotile(x, y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
            End If
        Next
    End With
End Sub

Public Sub DrawMapFringeTile(ByVal x As Long, ByVal y As Long)
Dim i As Long
    
    With Map.Tile(x, y)
        ' draw the map
        For i = MapLayer.Fringe To MapLayer.Fringe2
            ' skip tile if tileset isn't set
            If Autotile(x, y).Layer(i).renderState = RENDER_STATE_NORMAL Then
                ' Draw normally
                RenderTexture Tex_Tileset(.Layer(i).Tileset), ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), .Layer(i).x * 32, .Layer(i).y * 32, 32, 32, 32, 32
            ElseIf Autotile(x, y).Layer(i).renderState = RENDER_STATE_AUTOTILE Then
                ' Draw autotiles
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY(y * PIC_Y), 1, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY(y * PIC_Y), 2, x, y
                DrawAutoTile i, ConvertMapX(x * PIC_X), ConvertMapY((y * PIC_Y) + 16), 3, x, y
                DrawAutoTile i, ConvertMapX((x * PIC_X) + 16), ConvertMapY((y * PIC_Y) + 16), 4, x, y
            End If
        Next
    End With
End Sub

Public Sub DrawBars()
Dim left As Long, top As Long, Width As Long, height As Long
Dim tmpX As Long, tmpY As Long, barWidth As Long, i As Long, npcNum As Long
Dim partyIndex As Long

    ' dynamic bar calculations
    Width = D3DT_TEXTURE(Tex_Bars).Width
    height = D3DT_TEXTURE(Tex_Bars).height / 4
    
    ' render npc health bars
    For i = 1 To MAX_MAP_NPCS
        npcNum = MapNpc(i).num
        ' exists?
        If npcNum > 0 Then
            ' alive?
            If MapNpc(i).Vital(Vitals.HP) > 0 And MapNpc(i).Vital(Vitals.HP) < Npc(npcNum).HP Then
                ' lock to npc
                tmpX = MapNpc(i).x * PIC_X + MapNpc(i).xOffset + 16 - (Width / 2)
                tmpY = MapNpc(i).y * PIC_Y + MapNpc(i).yOffset + 35
                
                ' calculate the width to fill
                If Width > 0 Then BarWidth_NpcHP_Max(i) = ((MapNpc(i).Vital(Vitals.HP) / Width) / (Npc(npcNum).HP / Width)) * Width
                
                ' draw bar background
                top = height * 1 ' HP bar background
                left = 0
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, Width, height, Width, height
                
                ' draw the bar proper
                top = 0 ' HP bar
                left = 0
                RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, BarWidth_NpcHP(i), height, BarWidth_NpcHP(i), height
            End If
        End If
    Next

    ' check for casting time bar
    If SpellBuffer > 0 Then
        If Spell(PlayerSpells(SpellBuffer).Spell).CastTime > 0 Then
            ' lock to player
            tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (Width / 2)
            tmpY = GetPlayerY(MyIndex) * PIC_Y + Player(MyIndex).yOffset + 35 + height + 1
            
            ' calculate the width to fill
            If Width > 0 Then barWidth = (GetTickCount - SpellBufferTimer) / ((Spell(PlayerSpells(SpellBuffer).Spell).CastTime * 1000)) * Width
            
            ' draw bar background
            top = height * 3 ' cooldown bar background
            left = 0
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, Width, height, Width, height
             
            ' draw the bar proper
            top = height * 2 ' cooldown bar
            left = 0
            RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, barWidth, height, barWidth, height
        End If
    End If
    
    ' draw own health bar
    If GetPlayerVital(MyIndex, Vitals.HP) > 0 And GetPlayerVital(MyIndex, Vitals.HP) < GetPlayerMaxVital(MyIndex, Vitals.HP) Then
        ' lock to Player
        tmpX = GetPlayerX(MyIndex) * PIC_X + Player(MyIndex).xOffset + 16 - (Width / 2)
        tmpY = GetPlayerY(MyIndex) * PIC_X + Player(MyIndex).yOffset + 35
       
        ' calculate the width to fill
        If Width > 0 Then BarWidth_PlayerHP_Max(MyIndex) = ((GetPlayerVital(MyIndex, Vitals.HP) / Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / Width)) * Width
       
        ' draw bar background
        top = height * 1 ' HP bar background
        left = 0
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, Width, height, Width, height
       
        ' draw the bar proper
        top = 0 ' HP bar
        left = 0
        RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, BarWidth_PlayerHP(MyIndex), height, BarWidth_PlayerHP(MyIndex), height
    End If
    
    ' draw party health bars
    If Party.Leader > 0 Then
        For i = 1 To MAX_PARTY_MEMBERS
            partyIndex = Party.Member(i)
            If (partyIndex > 0) And (partyIndex <> MyIndex) And (GetPlayerMap(partyIndex) = GetPlayerMap(MyIndex)) Then
                ' player exists
                If GetPlayerVital(partyIndex, Vitals.HP) > 0 And GetPlayerVital(partyIndex, Vitals.HP) < GetPlayerMaxVital(partyIndex, Vitals.HP) Then
                    ' lock to Player
                    tmpX = GetPlayerX(partyIndex) * PIC_X + Player(partyIndex).xOffset + 16 - (Width / 2)
                    tmpY = GetPlayerY(partyIndex) * PIC_X + Player(partyIndex).yOffset + 35
                    
                    ' calculate the width to fill
                    BarWidth_PlayerHP_Max(partyIndex) = ((GetPlayerVital(partyIndex, Vitals.HP) / Width) / (GetPlayerMaxVital(partyIndex, Vitals.HP) / Width)) * Width
                    
                    ' draw bar background
                    top = height * 1 ' HP bar background
                    left = 0
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, Width, height, Width, height
                    
                    ' draw the bar proper
                    top = 0 ' HP bar
                    left = 0
                    RenderTexture Tex_Bars, ConvertMapX(tmpX), ConvertMapY(tmpY), left, top, BarWidth_PlayerHP(partyIndex), height, BarWidth_PlayerHP(partyIndex), height
                End If
            End If
        Next
    End If
End Sub

Public Sub DrawChatBubble(ByVal Index As Long)
Dim theArray() As String, x As Long, y As Long, i As Long, MaxWidth As Long, x2 As Long, y2 As Long, colour As Long
    
    With chatBubble(Index)
        If .targetType = TARGET_TYPE_PLAYER Then
            ' it's a player
            If GetPlayerMap(.target) = GetPlayerMap(MyIndex) Then
                ' change the colour depending on access
                colour = DarkBrown
                
                ' it's on our map - get co-ords
                x = ConvertMapX((Player(.target).x * 32) + Player(.target).xOffset) + 16
                y = ConvertMapY((Player(.target).y * 32) + Player(.target).yOffset) - 32
                
                ' word wrap the text
                WordWrap_Array .Msg, ChatBubbleWidth, theArray
                
                ' find max width
                For i = 1 To UBound(theArray)
                    If EngineGetTextWidth(Font_Default, theArray(i)) > MaxWidth Then MaxWidth = EngineGetTextWidth(Font_Default, theArray(i))
                Next
                
                ' calculate the new position
                x2 = x - (MaxWidth \ 2)
                y2 = y - (UBound(theArray) * 12)
                
                ' render bubble - top left
                RenderTexture Tex_GUI(37), x2 - 9, y2 - 5, 0, 0, 9, 5, 9, 5
                ' top right
                RenderTexture Tex_GUI(37), x2 + MaxWidth, y2 - 5, 119, 0, 9, 5, 9, 5
                ' top
                RenderTexture Tex_GUI(37), x2, y2 - 5, 9, 0, MaxWidth, 5, 5, 5
                ' bottom left
                RenderTexture Tex_GUI(37), x2 - 9, y, 0, 19, 9, 6, 9, 6
                ' bottom right
                RenderTexture Tex_GUI(37), x2 + MaxWidth, y, 119, 19, 9, 6, 9, 6
                ' bottom - left half
                RenderTexture Tex_GUI(37), x2, y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
                ' bottom - right half
                RenderTexture Tex_GUI(37), x2 + (MaxWidth \ 2) + 6, y, 9, 19, (MaxWidth \ 2) - 5, 6, 9, 6
                ' left
                RenderTexture Tex_GUI(37), x2 - 9, y2, 0, 6, 9, (UBound(theArray) * 12), 9, 1
                ' right
                RenderTexture Tex_GUI(37), x2 + MaxWidth, y2, 119, 6, 9, (UBound(theArray) * 12), 9, 1
                ' center
                RenderTexture Tex_GUI(37), x2, y2, 9, 5, MaxWidth, (UBound(theArray) * 12), 1, 1
                ' little pointy bit
                RenderTexture Tex_GUI(37), x - 5, y, 58, 19, 11, 11, 11, 11
                
                ' render each line centralised
                For i = 1 To UBound(theArray)
                    RenderText Font_Georgia, theArray(i), x - (EngineGetTextWidth(Font_Default, theArray(i)) / 2), y2, colour
                    y2 = y2 + 12
                Next
            End If
        End If
        ' check if it's timed out - close it if so
        If .timer + 5000 < GetTickCount Then
            .active = False
        End If
    End With
End Sub

Public Function isConstAnimated(ByVal sprite As Long) As Boolean
    isConstAnimated = False
    Select Case sprite
        Case 16, 21, 22, 26, 28
            isConstAnimated = True
    End Select
End Function

Public Function hasSpriteShadow(ByVal sprite As Long) As Boolean
    hasSpriteShadow = True
    Select Case sprite
        Case 25, 26
            hasSpriteShadow = False
    End Select
End Function

Public Sub DrawPlayer(ByVal Index As Long)
    Dim Anim As Byte
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim sprite As Long, spritetop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long
    
    ' pre-load sprite for calculations
    sprite = GetPlayerSprite(Index)
    'SetTexture Tex_Char(Sprite)

    If sprite < 1 Or sprite > Count_Char Then Exit Sub

    ' speed from weapon
    If GetPlayerEquipment(Index, Weapon) > 0 Then
        attackspeed = Item(GetPlayerEquipment(Index, Weapon)).Speed
    Else
        attackspeed = 1000
    End If

    If Not isConstAnimated(GetPlayerSprite(Index)) Then
        ' Reset frame
        Anim = 1
        ' Check for attacking animation
        If Player(Index).AttackTimer + (attackspeed / 2) > GetTickCount Then
            If Player(Index).Attacking = 1 Then
                Anim = 2
            End If
        Else
            ' If not attacking, walk normally
            Select Case GetPlayerDir(Index)
                Case DIR_UP
                    If (Player(Index).yOffset > 8) Then Anim = Player(Index).Step
                Case DIR_DOWN
                    If (Player(Index).yOffset < -8) Then Anim = Player(Index).Step
                Case DIR_LEFT
                    If (Player(Index).xOffset > 8) Then Anim = Player(Index).Step
                Case DIR_RIGHT
                    If (Player(Index).xOffset < -8) Then Anim = Player(Index).Step
            End Select
        End If
    Else
        If Player(Index).AnimTimer + 100 <= GetTickCount Then
            Player(Index).Anim = Player(Index).Anim + 1
            If Player(Index).Anim >= 3 Then Player(Index).Anim = 0
            Player(Index).AnimTimer = GetTickCount
        End If
        Anim = Player(Index).Anim
    End If

    ' Check to see if we want to stop making him attack
    With Player(Index)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

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
    End Select

    With rec
        .top = spritetop * (D3DT_TEXTURE(Tex_Char(sprite)).height / 4)
        .height = (D3DT_TEXTURE(Tex_Char(sprite)).height / 4)
        .left = Anim * (D3DT_TEXTURE(Tex_Char(sprite)).Width / 4)
        .Width = (D3DT_TEXTURE(Tex_Char(sprite)).Width / 4)
    End With

    ' Calculate the X
    x = GetPlayerX(Index) * PIC_X + Player(Index).xOffset - ((D3DT_TEXTURE(Tex_Char(sprite)).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (D3DT_TEXTURE(Tex_Char(sprite)).height) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - ((D3DT_TEXTURE(Tex_Char(sprite)).height / 4) - 32) - 4
    Else
        ' Proceed as normal
        y = GetPlayerY(Index) * PIC_Y + Player(Index).yOffset - 4
    End If
    
    RenderTexture Tex_Char(sprite), ConvertMapX(x), ConvertMapY(y), rec.left, rec.top, rec.Width, rec.height, rec.Width, rec.height
End Sub

Public Sub DrawNpc(ByVal MapNpcNum As Long)
    Dim Anim As Byte
    Dim i As Long
    Dim x As Long
    Dim y As Long
    Dim sprite As Long, spritetop As Long
    Dim rec As GeomRec
    Dim attackspeed As Long
    
    If MapNpc(MapNpcNum).num = 0 Then Exit Sub ' no npc set
    
    ' pre-load texture for calculations
    sprite = Npc(MapNpc(MapNpcNum).num).sprite
    'SetTexture Tex_Char(Sprite)

    If sprite < 1 Or sprite > Count_Char Then Exit Sub

    attackspeed = 1000

    If Not isConstAnimated(Npc(MapNpc(MapNpcNum).num).sprite) Then
        ' Reset frame
        Anim = 1
        ' Check for attacking animation
        If MapNpc(MapNpcNum).AttackTimer + (attackspeed / 2) > GetTickCount Then
            If MapNpc(MapNpcNum).Attacking = 1 Then
                Anim = 2
            End If
        Else
            ' If not attacking, walk normally
            Select Case MapNpc(MapNpcNum).dir
                Case DIR_UP
                    If (MapNpc(MapNpcNum).yOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_DOWN
                    If (MapNpc(MapNpcNum).yOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_LEFT
                    If (MapNpc(MapNpcNum).xOffset > 8) Then Anim = MapNpc(MapNpcNum).Step
                Case DIR_RIGHT
                    If (MapNpc(MapNpcNum).xOffset < -8) Then Anim = MapNpc(MapNpcNum).Step
            End Select
        End If
    Else
        With MapNpc(MapNpcNum)
            If .AnimTimer + 100 <= GetTickCount Then
                .Anim = .Anim + 1
                If .Anim >= 3 Then .Anim = 0
                .AnimTimer = GetTickCount
            End If
            Anim = .Anim
        End With
    End If

    ' Check to see if we want to stop making him attack
    With MapNpc(MapNpcNum)
        If .AttackTimer + attackspeed < GetTickCount Then
            .Attacking = 0
            .AttackTimer = 0
        End If
    End With

    ' Set the left
    Select Case MapNpc(MapNpcNum).dir
        Case DIR_UP
            spritetop = 3
        Case DIR_RIGHT
            spritetop = 2
        Case DIR_DOWN
            spritetop = 0
        Case DIR_LEFT
            spritetop = 1
    End Select

    With rec
        .top = (D3DT_TEXTURE(Tex_Char(sprite)).height / 4) * spritetop
        .height = D3DT_TEXTURE(Tex_Char(sprite)).height / 4
        .left = Anim * (D3DT_TEXTURE(Tex_Char(sprite)).Width / 4)
        .Width = (D3DT_TEXTURE(Tex_Char(sprite)).Width / 4)
    End With

    ' Calculate the X
    x = MapNpc(MapNpcNum).x * PIC_X + MapNpc(MapNpcNum).xOffset - ((D3DT_TEXTURE(Tex_Char(sprite)).Width / 4 - 32) / 2)

    ' Is the player's height more than 32..?
    If (D3DT_TEXTURE(Tex_Char(sprite)).height / 4) > 32 Then
        ' Create a 32 pixel offset for larger sprites
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - ((D3DT_TEXTURE(Tex_Char(sprite)).height / 4) - 32) - 4
    Else
        ' Proceed as normal
        y = MapNpc(MapNpcNum).y * PIC_Y + MapNpc(MapNpcNum).yOffset - 4
    End If
    
    RenderTexture Tex_Char(sprite), ConvertMapX(x), ConvertMapY(y), rec.left, rec.top, rec.Width, rec.height, rec.Width, rec.height
End Sub

Public Sub DrawShadow(ByVal sprite As Long, ByVal x As Long, ByVal y As Long)
    If hasSpriteShadow(sprite) Then RenderTexture Tex_Shadow, ConvertMapX(x), ConvertMapY(y), 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawTarget(ByVal x As Long, ByVal y As Long)
Dim Width As Long, height As Long
    
    ' calculations
    Width = D3DT_TEXTURE(Tex_Target).Width / 2
    height = D3DT_TEXTURE(Tex_Target).height
    
    x = x - ((Width - 32) / 2)
    y = y - (height / 2) + 16
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)

    'EngineRenderRectangle Tex_Target, x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Target, x, y, 0, 0, Width, height, Width, height
End Sub

Public Sub DrawTargetHover()
Dim i As Long, x As Long, y As Long, Width As Long, height As Long

    Width = D3DT_TEXTURE(Tex_Target).Width / 2
    height = D3DT_TEXTURE(Tex_Target).height
    
    If Width <= 0 Then Width = 1
    If height <= 0 Then height = 1
    
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(MyIndex) = GetPlayerMap(i) Then
            x = (Player(i).x * 32) + Player(i).xOffset + 32
            y = (Player(i).y * 32) + Player(i).yOffset + 32
            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    x = ConvertMapX(x)
                    y = ConvertMapY(y)
                    RenderTexture Tex_Target, x - 16 - (Width / 2), y - 16 - (height / 2), Width, 0, Width, height, Width, height
                End If
            End If
        End If
    Next
    
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            x = (MapNpc(i).x * 32) + MapNpc(i).xOffset + 32
            y = (MapNpc(i).y * 32) + MapNpc(i).yOffset + 32
            If x >= GlobalX_Map And x <= GlobalX_Map + 32 Then
                If y >= GlobalY_Map And y <= GlobalY_Map + 32 Then
                    x = ConvertMapX(x)
                    y = ConvertMapY(y)
                    RenderTexture Tex_Target, x - 16 - (Width / 2), y - 16 - (height / 2), Width, 0, Width, height, Width, height
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawResource(ByVal Resource_num As Long)
Dim Resource_master As Long
Dim Resource_state As Long
Dim Resource_sprite As Long
Dim rec As RECT
Dim x As Long, y As Long
Dim Width As Long, height As Long
    
    x = MapResource(Resource_num).x
    y = MapResource(Resource_num).y
    
    If x < 0 Or x > Map.MaxX Then Exit Sub
    If y < 0 Or y > Map.MaxY Then Exit Sub
    
    ' Get the Resource type
    Resource_master = Map.Tile(x, y).Data1
    
    If Resource_master = 0 Then Exit Sub

    If Resource(Resource_master).ResourceImage = 0 Then Exit Sub
    ' Get the Resource state
    Resource_state = MapResource(Resource_num).ResourceState

    If Resource_state = 0 Then ' normal
        Resource_sprite = Resource(Resource_master).ResourceImage
    ElseIf Resource_state = 1 Then ' used
        Resource_sprite = Resource(Resource_master).ExhaustedImage
    End If
    
    ' pre-load texture for calculations
    'SetTexture Tex_Resource(Resource_sprite)

    ' src rect
    With rec
        .top = 0
        .bottom = D3DT_TEXTURE(Tex_Resource(Resource_sprite)).height
        .left = 0
        .Right = D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Width
    End With

    ' Set base x + y, then the offset due to size
    x = (MapResource(Resource_num).x * PIC_X) - (D3DT_TEXTURE(Tex_Resource(Resource_sprite)).Width / 2) + 16
    y = (MapResource(Resource_num).y * PIC_Y) - D3DT_TEXTURE(Tex_Resource(Resource_sprite)).height + 32
    
    Width = rec.Right - rec.left
    height = rec.bottom - rec.top
    'EngineRenderRectangle Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, width, height, width, height, width, height
    RenderTexture Tex_Resource(Resource_sprite), ConvertMapX(x), ConvertMapY(y), 0, 0, Width, height, Width, height
End Sub

Public Sub DrawItem(ByVal itemnum As Long)
Dim PicNum As Integer, dontRender As Boolean, i As Long, tmpIndex As Long
    
    PicNum = Item(MapItem(itemnum).num).Pic

    If PicNum < 1 Or PicNum > Count_Item Then Exit Sub

     ' if it's not us then don't render
    If MapItem(itemnum).playerName <> vbNullString Then
        If Trim$(MapItem(itemnum).playerName) <> Trim$(GetPlayerName(MyIndex)) Then
            dontRender = True
        End If
        ' make sure it's not a party drop
        If Party.Leader > 0 Then
            For i = 1 To MAX_PARTY_MEMBERS
                tmpIndex = Party.Member(i)
                If tmpIndex > 0 Then
                    If Trim$(GetPlayerName(tmpIndex)) = Trim$(MapItem(itemnum).playerName) Then
                        If MapItem(itemnum).bound = 0 Then
                            dontRender = False
                        End If
                    End If
                End If
            Next
        End If
    End If
    
    'If Not dontRender Then EngineRenderRectangle Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32, 32, 32
    If Not dontRender Then
        RenderTexture Tex_Item(PicNum), ConvertMapX(MapItem(itemnum).x * PIC_X), ConvertMapY(MapItem(itemnum).y * PIC_Y), 0, 0, 32, 32, 32, 32
    End If
End Sub

Public Sub DrawDragItem()
    Dim PicNum As Integer, itemnum As Long
    
    If DragInvSlotNum = 0 Then Exit Sub
    
    itemnum = GetPlayerInvItemNum(MyIndex, DragInvSlotNum)
    If Not itemnum > 0 Then Exit Sub
    
    PicNum = Item(itemnum).Pic

    If PicNum < 1 Or PicNum > Count_Item Then Exit Sub

    'EngineRenderRectangle Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawDragSpell()
    Dim PicNum As Integer, spellnum As Long
    
    If DragSpell = 0 Then Exit Sub
    
    spellnum = PlayerSpells(DragSpell).Spell
    If Not spellnum > 0 Then Exit Sub
    
    PicNum = Spell(spellnum).Icon

    If PicNum < 1 Or PicNum > Count_Spellicon Then Exit Sub

    'EngineRenderRectangle Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Spellicon(PicNum), GlobalX - 16, GlobalY - 16, 0, 0, 32, 32, 32, 32
End Sub

Public Sub DrawAnimation(ByVal Index As Long, ByVal Layer As Long)
Dim sprite As Integer, sRECT As GeomRec, i As Long, Width As Long, height As Long, looptime As Long, FrameCount As Long
Dim x As Long, y As Long, lockindex As Long
    
    If AnimInstance(Index).Animation = 0 Then
        ClearAnimInstance Index
        Exit Sub
    End If
    
    sprite = Animation(AnimInstance(Index).Animation).sprite(Layer)
    
    If sprite < 1 Or sprite > Count_Anim Then Exit Sub
    
    ' pre-load texture for calculations
    'SetTexture Tex_Anim(Sprite)
    
    FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
    
    ' total width divided by frame count
    Width = 192 'D3DT_TEXTURE(Tex_Anim(Sprite)).width / frameCount
    height = 192 'D3DT_TEXTURE(Tex_Anim(Sprite)).height
    
    With sRECT
        .top = (height * ((AnimInstance(Index).FrameIndex(Layer) - 1) \ AnimColumns))
        .height = height
        .left = (Width * (((AnimInstance(Index).FrameIndex(Layer) - 1) Mod AnimColumns)))
        .Width = Width
    End With
    
    ' change x or y if locked
    If AnimInstance(Index).LockType > TARGET_TYPE_NONE Then ' if <> none
        ' is a player
        If AnimInstance(Index).LockType = TARGET_TYPE_PLAYER Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if is ingame
            If IsPlaying(lockindex) Then
                ' check if on same map
                If GetPlayerMap(lockindex) = GetPlayerMap(MyIndex) Then
                    ' is on map, is playing, set x & y
                    x = (GetPlayerX(lockindex) * PIC_X) + 16 - (Width / 2) + Player(lockindex).xOffset
                    y = (GetPlayerY(lockindex) * PIC_Y) + 16 - (height / 2) + Player(lockindex).yOffset
                End If
            End If
        ElseIf AnimInstance(Index).LockType = TARGET_TYPE_NPC Then
            ' quick save the index
            lockindex = AnimInstance(Index).lockindex
            ' check if NPC exists
            If MapNpc(lockindex).num > 0 Then
                ' check if alive
                If MapNpc(lockindex).Vital(Vitals.HP) > 0 Then
                    ' exists, is alive, set x & y
                    x = (MapNpc(lockindex).x * PIC_X) + 16 - (Width / 2) + MapNpc(lockindex).xOffset
                    y = (MapNpc(lockindex).y * PIC_Y) + 16 - (height / 2) + MapNpc(lockindex).yOffset
                Else
                    ' npc not alive anymore, kill the animation
                    ClearAnimInstance Index
                    Exit Sub
                End If
            Else
                ' npc not alive anymore, kill the animation
                ClearAnimInstance Index
                Exit Sub
            End If
        End If
    Else
        ' no lock, default x + y
        x = (AnimInstance(Index).x * 32) + 16 - (Width / 2)
        y = (AnimInstance(Index).y * 32) + 16 - (height / 2)
    End If
    
    x = ConvertMapX(x)
    y = ConvertMapY(y)
    
    'EngineRenderRectangle Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.width, sRECT.height, sRECT.width, sRECT.height, sRECT.width, sRECT.height
    RenderTexture Tex_Anim(sprite), x, y, sRECT.left, sRECT.top, sRECT.Width, sRECT.height, sRECT.Width, sRECT.height
End Sub

Public Sub DrawInventoryItemDesc()
Dim invSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_INVENTORY).visible Then Exit Sub
    If DragInvSlotNum > 0 Then Exit Sub
    
    invSlot = IsInvItem(GlobalX, GlobalY)
    If invSlot > 0 Then
        If GetPlayerInvItemNum(MyIndex, invSlot) > 0 Then
            If Item(GetPlayerInvItemNum(MyIndex, invSlot)).BindType > 0 And PlayerInv(invSlot).bound > 0 Then isSB = True
            DrawItemDesc GetPlayerInvItemNum(MyIndex, invSlot), GUIWindow(GUI_INVENTORY).x - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).y, isSB
            ' value
            If InShop > 0 Then
                DrawItemCost False, invSlot, GUIWindow(GUI_INVENTORY).x - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_INVENTORY).y + GUIWindow(GUI_DESCRIPTION).height + 10
            End If
        End If
    End If
End Sub

Public Sub DrawShopItemDesc()
Dim shopSlot As Long
    
    If Not GUIWindow(GUI_SHOP).visible Then Exit Sub
    
    shopSlot = IsShopItem(GlobalX, GlobalY)
    If shopSlot > 0 Then
        If Shop(InShop).TradeItem(shopSlot).Item > 0 Then
            DrawItemDesc Shop(InShop).TradeItem(shopSlot).Item, GUIWindow(GUI_SHOP).x + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).y
            DrawItemCost True, shopSlot, GUIWindow(GUI_SHOP).x + GUIWindow(GUI_SHOP).Width + 10, GUIWindow(GUI_SHOP).y + GUIWindow(GUI_DESCRIPTION).height + 10
        End If
    End If
End Sub

Public Sub DrawCharacterItemDesc()
Dim eqSlot As Long, isSB As Boolean
    
    If Not GUIWindow(GUI_CHARACTER).visible Then Exit Sub
    
    eqSlot = IsEqItem(GlobalX, GlobalY)
    If eqSlot > 0 Then
        If GetPlayerEquipment(MyIndex, eqSlot) > 0 Then
            If Item(GetPlayerEquipment(MyIndex, eqSlot)).BindType > 0 Then isSB = True
            DrawItemDesc GetPlayerEquipment(MyIndex, eqSlot), GUIWindow(GUI_CHARACTER).x - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_CHARACTER).y, isSB
        End If
    End If
End Sub

Public Sub DrawItemCost(ByVal isShop As Boolean, ByVal slotNum As Long, ByVal x As Long, ByVal y As Long)
Dim CostItem As Long, CostValue As Long, itemnum As Long, sString As String, Width As Long, height As Long

    If slotNum = 0 Then Exit Sub
    
    If InShop <= 0 Then Exit Sub
    
    ' draw the window
    Width = 190
    height = 36

    'EngineRenderRectangle Tex_GUI(33), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(33), x, y, 0, 0, Width, height, Width, height
    
    ' find out the cost
    If Not isShop Then
        ' inventory - default to gold
        itemnum = GetPlayerInvItemNum(MyIndex, slotNum)
        If itemnum = 0 Then Exit Sub
        CostItem = 1
        CostValue = (Item(itemnum).Price / 100) * Shop(InShop).BuyRate
        sString = "The shop will buy for"
    Else
        itemnum = Shop(InShop).TradeItem(slotNum).Item
        If itemnum = 0 Then Exit Sub
        CostItem = Shop(InShop).TradeItem(slotNum).CostItem
        CostValue = Shop(InShop).TradeItem(slotNum).CostValue
        sString = "The shop will sell for"
    End If
    
    'EngineRenderRectangle Tex_Item(Item(CostItem).Pic), x + 155, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Item(Item(CostItem).Pic), x + 155, y + 2, 0, 0, 32, 32, 32, 32
    
    RenderText Font_Default, sString, x + 4, y + 3, DarkGrey
    
    RenderText Font_Default, CostValue & " " & Trim$(Item(CostItem).Name), x + 4, y + 18, White
End Sub

Public Sub DrawItemDesc(ByVal itemnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal soulBound As Boolean = False)
Dim colour As Long, descString As String, theName As String, className As String, levelTxt As String, sInfo() As String, i As Long, Width As Long, height As Long
    
    ' get out
    If itemnum = 0 Then Exit Sub

    ' render the window
    Width = 190
    height = 126
    'EngineRenderRectangle Tex_GUI(8), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(8), x, y, 0, 0, Width, height, Width, height
    
    ' make sure it has a sprite
    If Item(itemnum).Pic > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Item(Item(itemnum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 64, 64
        RenderTexture Tex_Item(Item(itemnum).Pic), x + 16, y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    ' work out name colour
    Select Case Item(itemnum).Rarity
        Case 0 ' white
            colour = White
        Case 1 ' green
            colour = Green
        Case 2 ' blue
            colour = Blue
        Case 3 ' maroon
            colour = Red
        Case 4 ' purple
            colour = Pink
        Case 5 ' orange
            colour = Brown
    End Select
    
    If Not soulBound Then
        theName = Trim$(Item(itemnum).Name)
    Else
        theName = "(SB) " & Trim$(Item(itemnum).Name)
    End If
    
    ' render name
    RenderText Font_Default, theName, x + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), y + 6, colour
    
    ' class req
    If Item(itemnum).ClassReq > 0 Then
        className = Trim$(Class(Item(itemnum).ClassReq).Name)
        ' do we match it?
        If GetPlayerClass(MyIndex) = Item(itemnum).ClassReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        className = "No class req."
        colour = Green
    End If
    RenderText Font_Default, className, x + 48 - (EngineGetTextWidth(Font_Default, className) \ 2), y + 92, colour
    
    ' level
    If Item(itemnum).LevelReq > 0 Then
        levelTxt = "Level " & Item(itemnum).LevelReq
        ' do we match it?
        If GetPlayerLevel(MyIndex) >= Item(itemnum).LevelReq Then
            colour = Green
        Else
            colour = BrightRed
        End If
    Else
        levelTxt = "No level req."
        colour = Green
    End If
    RenderText Font_Default, levelTxt, x + 48 - (EngineGetTextWidth(Font_Default, levelTxt) \ 2), y + 107, colour
    
    ' first we cache all information strings then loop through and render them

    ' item type
    i = 1
    ReDim Preserve sInfo(1 To i) As String
    Select Case Item(itemnum).Type
        Case ITEM_TYPE_NONE
            sInfo(i) = "No type"
        Case ITEM_TYPE_WEAPON
            sInfo(i) = "Weapon"
        Case ITEM_TYPE_ARMOR
            sInfo(i) = "Armour"
        Case ITEM_TYPE_HELMET
            sInfo(i) = "Helmet"
        Case ITEM_TYPE_SHIELD
            sInfo(i) = "Shield"
        Case ITEM_TYPE_CONSUME
            sInfo(i) = "Consume"
        Case ITEM_TYPE_KEY
            sInfo(i) = "Key"
        Case ITEM_TYPE_CURRENCY
            sInfo(i) = "Currency"
        Case ITEM_TYPE_SPELL
            sInfo(i) = "Spell"
    End Select
    
    ' more info
    Select Case Item(itemnum).Type
        Case ITEM_TYPE_NONE, ITEM_TYPE_KEY, ITEM_TYPE_CURRENCY
            ' binding
            If Item(itemnum).BindType = 1 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Bind on Pickup"
            ElseIf Item(itemnum).BindType = 2 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Bind on Equip"
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Value: " & Item(itemnum).Price & "g"
        Case ITEM_TYPE_WEAPON, ITEM_TYPE_ARMOR, ITEM_TYPE_HELMET, ITEM_TYPE_SHIELD
            ' damage/defence
            If Item(itemnum).Type = ITEM_TYPE_WEAPON Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Damage: " & Item(itemnum).Data2
                ' speed
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Speed: " & (Item(itemnum).Speed / 1000) & "s"
            Else
                If Item(itemnum).Data2 > 0 Then
                    i = i + 1
                    ReDim Preserve sInfo(1 To i) As String
                    sInfo(i) = "Defence: " & Item(itemnum).Data2
                End If
            End If
            ' binding
            If Item(itemnum).BindType = 1 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Bind on Pickup"
            ElseIf Item(itemnum).BindType = 2 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Bind on Equip"
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Value: " & Item(itemnum).Price & "g"
            ' stat bonuses
            If Item(itemnum).Add_Stat(Stats.Strength) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemnum).Add_Stat(Stats.Strength) & " Str"
            End If
            If Item(itemnum).Add_Stat(Stats.Endurance) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemnum).Add_Stat(Stats.Endurance) & " End"
            End If
            If Item(itemnum).Add_Stat(Stats.Intelligence) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemnum).Add_Stat(Stats.Intelligence) & " Int"
            End If
            If Item(itemnum).Add_Stat(Stats.Agility) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemnum).Add_Stat(Stats.Agility) & " Agi"
            End If
            If Item(itemnum).Add_Stat(Stats.Willpower) > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemnum).Add_Stat(Stats.Willpower) & " Will"
            End If
        Case ITEM_TYPE_CONSUME
            If Item(itemnum).CastSpell > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Casts Spell"
            End If
            If Item(itemnum).AddHP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemnum).AddHP & " HP"
            End If
            If Item(itemnum).AddMP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemnum).AddMP & " SP"
            End If
            If Item(itemnum).AddEXP > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "+" & Item(itemnum).AddEXP & " EXP"
            End If
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Value: " & Item(itemnum).Price & "g"
        Case ITEM_TYPE_SPELL
            ' price
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Value: " & Item(itemnum).Price & "g"
    End Select
    
    ' go through and render all this shit
    y = y + 12
    For i = 1 To UBound(sInfo)
        y = y + 12
        RenderText Font_Default, sInfo(i), x + 141 - (EngineGetTextWidth(Font_Default, sInfo(i)) \ 2), y, White
    Next
End Sub

Public Sub DrawInventory()
Dim i As Long, x As Long, y As Long, itemnum As Long, itempic As Long
Dim Amount As String
Dim colour As Long
Dim top As Long, left As Long
Dim Width As Long, height As Long

    ' render the window
    Width = 195
    height = 250
    'EngineRenderRectangle Tex_GUI(5), GUIWindow(GUI_INVENTORY).x, GUIWindow(GUI_INVENTORY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_INVENTORY).x, GUIWindow(GUI_INVENTORY).y, 0, 0, Width, height, Width, height
    
    For i = 1 To MAX_INV
        itemnum = GetPlayerInvItemNum(MyIndex, i)
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            
            ' exit out if we're offering item in a trade.
            If InTrade > 0 Then
                For x = 1 To MAX_INV
                    If TradeYourOffer(x).num = i Then
                        GoTo NextLoop
                    End If
                Next
            End If
            
            ' exit out if dragging
            If DragInvSlotNum = i Then GoTo NextLoop

            If itempic > 0 And itempic <= Count_Item Then
                top = GUIWindow(GUI_INVENTORY).y + InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                left = GUIWindow(GUI_INVENTORY).x + InvLeft + ((InvOffsetX + 32) * (((i - 1) Mod InvColumns)))

                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32

                ' If item is a stack - draw the amount you have
                If GetPlayerInvItemValue(MyIndex, i) > 1 Then
                    y = top + 21
                    x = left - 4
                    Amount = CStr(GetPlayerInvItemValue(MyIndex, i))
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), x, y, colour
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawPlayerSpellDesc()
Dim spellSlot As Long
    
    If Not GUIWindow(GUI_SPELLS).visible Then Exit Sub
    If DragSpell > 0 Then Exit Sub
    
    spellSlot = IsPlayerSpell(GlobalX, GlobalY)
    If spellSlot > 0 Then
        If PlayerSpells(spellSlot).Spell > 0 Then
            DrawSpellDesc PlayerSpells(spellSlot).Spell, GUIWindow(GUI_SPELLS).x - GUIWindow(GUI_DESCRIPTION).Width - 10, GUIWindow(GUI_SPELLS).y, spellSlot
        End If
    End If
End Sub

Public Sub DrawSpellDesc(ByVal spellnum As Long, ByVal x As Long, ByVal y As Long, Optional ByVal spellSlot As Long = 0)
Dim colour As Long, theName As String, sUse As String, sInfo() As String, i As Long, tmpWidth As Long, barWidth As Long
Dim Width As Long, height As Long
    
    ' don't show desc when dragging
    If DragSpell > 0 Then Exit Sub
    
    ' get out
    If spellnum = 0 Then Exit Sub

    ' render the window
    Width = 190
    height = 126
    'EngineRenderRectangle Tex_GUI(34), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(34), x, y, 0, 0, Width, height, Width, height
    
    ' make sure it has a sprite
    If Spell(spellnum).Icon > 0 Then
        ' render sprite
        'EngineRenderRectangle Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32, 32, 32
        RenderTexture Tex_Spellicon(Spell(spellnum).Icon), x + 16, y + 27, 0, 0, 64, 64, 32, 32
    End If
    
    ' render name
    colour = White
    theName = Trim$(Spell(spellnum).Name)
    RenderText Font_Default, theName, x + 95 - (EngineGetTextWidth(Font_Default, theName) \ 2), y + 6, colour
    
    ' if it's a player spell then do the rank up message
    colour = White
    If spellSlot > 0 Then
        ' draw the rank bar
        barWidth = 78
        If Spell(spellnum).NextRank > 0 Then
            tmpWidth = ((PlayerSpells(spellSlot).Uses / barWidth) / (Spell(spellnum).NextUses / barWidth)) * barWidth
        Else
            tmpWidth = 78
        End If
        'EngineRenderRectangle Tex_GUI(35), x + 9, y + 99, 0, 0, tmpWidth, 16, tmpWidth, 16, tmpWidth, 16
        RenderTexture Tex_GUI(35), x + 9, y + 99, 0, 0, tmpWidth, 16, tmpWidth, 16
        ' does it rank up?
        If Spell(spellnum).NextRank > 0 Then
            sUse = "Uses: " & PlayerSpells(spellSlot).Uses & "/" & Spell(spellnum).NextUses
            If PlayerSpells(spellSlot).Uses = Spell(spellnum).NextUses Then
                If Not GetPlayerLevel(MyIndex) >= Spell(Spell(spellnum).NextRank).LevelReq Then
                    colour = BrightRed
                    sUse = "Lvl " & Spell(Spell(spellnum).NextRank).LevelReq & " req."
                End If
            End If
        Else
            colour = Grey
            sUse = "Max Rank"
        End If
        RenderText Font_Default, sUse, x + 48 - (EngineGetTextWidth(Font_Default, sUse) \ 2), y + 99, colour
    End If
    
    ' first we cache all information strings then loop through and render them

    ' item type
    i = 1
    ReDim Preserve sInfo(1 To i) As String
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP
            sInfo(i) = "Damage HP"
        Case SPELL_TYPE_DAMAGEMP
            sInfo(i) = "Damage SP"
        Case SPELL_TYPE_HEALHP
            sInfo(i) = "Heal HP"
        Case SPELL_TYPE_HEALMP
            sInfo(i) = "Heal SP"
        Case SPELL_TYPE_WARP
            sInfo(i) = "Warp"
    End Select
    
    ' more info
    Select Case Spell(spellnum).Type
        Case SPELL_TYPE_DAMAGEHP, SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP
            ' damage
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Vital: " & Spell(spellnum).Vital
            
            ' mp cost
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cost: " & Spell(spellnum).MPCost & " SP"
            
            ' cast time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cast Time: " & Spell(spellnum).CastTime & "s"
            
            ' cd time
            i = i + 1
            ReDim Preserve sInfo(1 To i) As String
            sInfo(i) = "Cooldown: " & Spell(spellnum).CDTime & "s"
            
            ' aoe
            If Spell(spellnum).AoE > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "AoE: " & Spell(spellnum).AoE
            End If
            
            ' stun
            If Spell(spellnum).StunDuration > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "Stun: " & Spell(spellnum).StunDuration & "s"
            End If
            
            ' dot
            If Spell(spellnum).Duration > 0 And Spell(spellnum).Interval > 0 Then
                i = i + 1
                ReDim Preserve sInfo(1 To i) As String
                sInfo(i) = "DoT: " & (Spell(spellnum).Duration / Spell(spellnum).Interval) & " tick"
            End If
    End Select
    
    ' go through and render all this shit
    y = y + 12
    For i = 1 To UBound(sInfo)
        y = y + 12
        RenderText Font_Default, sInfo(i), x + 141 - (EngineGetTextWidth(Font_Default, sInfo(i)) \ 2), y, White
    Next
End Sub

Public Sub DrawSkills()
Dim i As Long, x As Long, y As Long, spellnum As Long, spellpic As Long
Dim top As Long, left As Long
Dim Width As Long, height As Long

    ' render the window
    Width = 195
    height = 250
    'EngineRenderRectangle Tex_GUI(5), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(5), GUIWindow(GUI_SPELLS).x, GUIWindow(GUI_SPELLS).y, 0, 0, Width, height, Width, height
    
    ' render skills
    For i = 1 To MAX_PLAYER_SPELLS
        spellnum = PlayerSpells(i).Spell

        ' make sure not dragging it
        If DragSpell = i Then GoTo NextLoop
        
        ' actually render
        If spellnum > 0 And spellnum <= MAX_SPELLS Then
            spellpic = Spell(spellnum).Icon

            If spellpic > 0 And spellpic <= Count_Spellicon Then
                top = GUIWindow(GUI_SPELLS).y + SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                left = GUIWindow(GUI_SPELLS).x + SpellLeft + ((SpellOffsetX + 32) * (((i - 1) Mod SpellColumns)))
                If SpellCD(i) > 0 Then
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                    RenderTexture Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                Else
                    'EngineRenderRectangle Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                    RenderTexture Tex_Spellicon(spellpic), left, top, 0, 0, 32, 32, 32, 32
                End If
            End If
        End If
NextLoop:
    Next
End Sub

Public Sub DrawEquipment()
Dim x As Long, y As Long, i As Long
Dim itemnum As Long, itempic As Long

    For i = 1 To Equipment.Equipment_Count - 1
        itemnum = GetPlayerEquipment(MyIndex, i)

        ' get the item sprite
        If itemnum > 0 Then
            itempic = Tex_Item(Item(itemnum).Pic)
        Else
            ' no item equiped - use blank image
            itempic = Tex_GUI(8 + i)
        End If
        
        y = GUIWindow(GUI_CHARACTER).y + EqTop
        x = GUIWindow(GUI_CHARACTER).x + EqLeft + ((EqOffsetX + 32) * (((i - 1) Mod EqColumns)))

        'EngineRenderRectangle itempic, x, y, 0, 0, 32, 32, 32, 32, 32, 32
        RenderTexture itempic, x, y, 0, 0, 32, 32, 32, 32
    Next
End Sub

Public Sub DrawCharacter()
Dim x As Long, y As Long, i As Long, dX As Long, dY As Long, tmpString As String, buttonnum As Long
Dim Width As Long, height As Long
    
    x = GUIWindow(GUI_CHARACTER).x
    y = GUIWindow(GUI_CHARACTER).y
    
    ' render the window
    Width = 195
    height = 250
    'EngineRenderRectangle Tex_GUI(6), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(6), x, y, 0, 0, Width, height, Width, height
    
    ' render name
    tmpString = Trim$(GetPlayerName(MyIndex)) & " - Level " & GetPlayerLevel(MyIndex)
    RenderText Font_Default, tmpString, x + 7 + (187 / 2) - (EngineGetTextWidth(Font_Default, tmpString) / 2), y + 9, White
    
    ' render stats
    dX = x + 20
    dY = y + 145
    RenderText Font_Default, "Str: " & GetPlayerStat(MyIndex, Strength), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "End: " & GetPlayerStat(MyIndex, Endurance), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Int: " & GetPlayerStat(MyIndex, Intelligence), dX, dY, White
    dY = y + 145
    dX = dX + 80
    RenderText Font_Default, "Agi: " & GetPlayerStat(MyIndex, Agility), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Will: " & GetPlayerStat(MyIndex, Willpower), dX, dY, White
    dY = dY + 15
    RenderText Font_Default, "Pnts: " & GetPlayerPOINTS(MyIndex), dX, dY, White
    
    ' draw the face
    If GetPlayerSprite(MyIndex) > 0 And GetPlayerSprite(MyIndex) <= Count_Face Then
        'EngineRenderRectangle Tex_Face(GetPlayerSprite(MyIndex)), x + 49, y + 38, 0, 0, 96, 96, 96, 96, 96, 96
        RenderTexture Tex_Face(GetPlayerSprite(MyIndex)), x + 49, y + 38, 0, 0, 96, 96, 96, 96
    End If
    
    ' draw the equipment
    DrawEquipment
    
    If GetPlayerPOINTS(MyIndex) > 0 Then
        ' draw the buttons
        For buttonnum = 16 To 20
            x = GUIWindow(GUI_CHARACTER).x + Buttons(buttonnum).x
            y = GUIWindow(GUI_CHARACTER).y + Buttons(buttonnum).y
            Width = Buttons(buttonnum).Width
            height = Buttons(buttonnum).height
            ' render accept button
            If Buttons(buttonnum).state = 2 Then
                ' we're clicked boyo
                Width = Buttons(buttonnum).Width
                height = Buttons(buttonnum).height
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
            ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
                ' play sound if needed
                If Not lastButtonSound = buttonnum Then
                    Play_Sound Sound_ButtonHover
                    lastButtonSound = buttonnum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
    End If
End Sub

Public Sub DrawOptions()
Dim i As Long, x As Long, y As Long
Dim Width As Long, height As Long

    ' render the window
    Width = 195
    height = 250
    'EngineRenderRectangle Tex_GUI(29), GUIWindow(GUI_OPTIONS).x, GUIWindow(GUI_OPTIONS).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(29), GUIWindow(GUI_OPTIONS).x, GUIWindow(GUI_OPTIONS).y, 0, 0, Width, height, Width, height
    
    ' draw buttons
    For i = 26 To 33
        ' set co-ordinate
        x = GUIWindow(GUI_OPTIONS).x + Buttons(i).x
        y = GUIWindow(GUI_OPTIONS).y + Buttons(i).y
        Width = Buttons(i).Width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ' play sound if needed
            If Not lastButtonSound = i Then
                Play_Sound Sound_ButtonHover
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawParty()
Dim i As Long, x As Long, y As Long, Width As Long, playerNum As Long, theName As String
Dim height As Long

    ' render the window
    Width = 195
    height = 250
    'EngineRenderRectangle Tex_GUI(7), GUIWindow(GUI_PARTY).x, GUIWindow(GUI_PARTY).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(7), GUIWindow(GUI_PARTY).x, GUIWindow(GUI_PARTY).y, 0, 0, Width, height, Width, height
    
    ' draw the bars
    If Party.Leader > 0 Then ' make sure we're in a party
        ' draw leader
        playerNum = Party.Leader
        ' name
        theName = Trim$(GetPlayerName(playerNum))
        ' draw name
        y = GUIWindow(GUI_PARTY).y + 12
        x = GUIWindow(GUI_PARTY).x + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
        RenderText Font_Default, theName, x, y, White
        ' draw hp
        y = GUIWindow(GUI_PARTY).y + 29
        x = GUIWindow(GUI_PARTY).x + 6
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
        End If
        'EngineRenderRectangle Tex_GUI(16), x, y, 0, 0, width, 9, width, 9, width, 9
        RenderTexture Tex_GUI(16), x, y, 0, 0, Width, 9, Width, 9
        ' draw mp
        y = GUIWindow(GUI_PARTY).y + 38
        ' make sure we actually have the data before rendering
        If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
            Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
        End If
        'EngineRenderRectangle Tex_GUI(17), x, y, 0, 0, width, 9, width, 9, width, 9
        RenderTexture Tex_GUI(17), x, y, 0, 0, Width, 9, Width, 9
        
        ' draw members
        For i = 1 To MAX_PARTY_MEMBERS
            If Party.Member(i) > 0 Then
                If Party.Member(i) <> Party.Leader Then
                    ' cache the index
                    playerNum = Party.Member(i)
                    ' name
                    theName = Trim$(GetPlayerName(playerNum))
                    ' draw name
                    y = GUIWindow(GUI_PARTY).y + 12 + ((i - 1) * 49)
                    x = GUIWindow(GUI_PARTY).x + 7 + 90 - (EngineGetTextWidth(Font_Default, theName) / 2)
                    RenderText Font_Default, theName, x, y, White
                    ' draw hp
                    y = GUIWindow(GUI_PARTY).y + 29 + ((i - 1) * 49)
                    x = GUIWindow(GUI_PARTY).x + 6
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.HP) > 0 And GetPlayerMaxVital(playerNum, Vitals.HP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(playerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(16), x, y, 0, 0, width, 9, width, 9, width, 9
                    RenderTexture Tex_GUI(16), x, y, 0, 0, Width, 9, Width, 9
                    ' draw mp
                    y = GUIWindow(GUI_PARTY).y + 38 + ((i - 1) * 49)
                    ' make sure we actually have the data before rendering
                    If GetPlayerVital(playerNum, Vitals.MP) > 0 And GetPlayerMaxVital(playerNum, Vitals.MP) > 0 Then
                        Width = ((GetPlayerVital(playerNum, Vitals.MP) / Party_SPRWidth) / (GetPlayerMaxVital(playerNum, Vitals.MP) / Party_SPRWidth)) * Party_SPRWidth
                    End If
                    'EngineRenderRectangle Tex_GUI(17), x, y, 0, 0, width, 9, width, 9, width, 9
                    RenderTexture Tex_GUI(17), x, y, 0, 0, Width, 9, Width, 9
                End If
            End If
        Next
    End If
    
    ' draw buttons
    For i = 24 To 25
        ' set co-ordinate
        x = GUIWindow(GUI_PARTY).x + Buttons(i).x
        y = GUIWindow(GUI_PARTY).y + Buttons(i).y
        Width = Buttons(i).Width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ' play sound if needed
            If Not lastButtonSound = i Then
                Play_Sound Sound_ButtonHover
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawHotbar()
Dim i As Long, x As Long, y As Long, t As Long, sS As String
Dim Width As Long, height As Long

    For i = 1 To MAX_HOTBAR
        ' draw the box
        x = GUIWindow(GUI_HOTBAR).x + ((i - 1) * (5 + 36))
        y = GUIWindow(GUI_HOTBAR).y
        Width = 36
        height = 36
        'EngineRenderRectangle Tex_GUI(2), x, y, 0, 0, width, height, width, height, width, heigh
        RenderTexture Tex_GUI(2), x, y, 0, 0, Width, height, Width, height
        ' draw the icon
        Select Case Hotbar(i).sType
            Case 1 ' inventory
                If Len(Item(Hotbar(i).Slot).Name) > 0 Then
                    If Item(Hotbar(i).Slot).Pic > 0 Then
                        'EngineRenderRectangle Tex_Item(Item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_Item(Item(Hotbar(i).Slot).Pic), x + 2, y + 2, 0, 0, 32, 32, 32, 32
                    End If
                End If
            Case 2 ' spell
                If Len(Spell(Hotbar(i).Slot).Name) > 0 Then
                    If Spell(Hotbar(i).Slot).Icon > 0 Then
                        ' render normal icon
                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32
                        RenderTexture Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32
                        ' we got the spell?
                        For t = 1 To MAX_PLAYER_SPELLS
                            If PlayerSpells(t).Spell > 0 Then
                                If PlayerSpells(t).Spell = Hotbar(i).Slot Then
                                    If SpellCD(t) > 0 Then
                                        'EngineRenderRectangle Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, 32, 32, , , , , , , 254, 190, 190, 190
                                        RenderTexture Tex_Spellicon(Spell(Hotbar(i).Slot).Icon), x + 2, y + 2, 0, 0, 32, 32, 32, 32, D3DColorARGB(255, 100, 100, 100)
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
        End Select
        ' draw the numbers
        sS = Str(i)
        If i = 10 Then sS = "0"
        If i = 11 Then sS = " -"
        If i = 12 Then sS = " ="
        RenderText Font_Default, sS, x + 4, y + 20, White
    Next
End Sub

Public Sub DrawGUI()
Dim i As Long, x As Long, y As Long
Dim Width As Long, height As Long

    ' render shadow
    'EngineRenderRectangle Tex_GUI(32), 0, 0, 0, 0, 800, 64, 1, 64, 800, 64
    'EngineRenderRectangle Tex_GUI(31), 0, 600 - 64, 0, 0, 800, 64, 1, 64, 800, 64
    RenderTexture Tex_GUI(32), 0, 0, 0, 0, 800, 64, 1, 64
    RenderTexture Tex_GUI(31), 0, 600 - 64, 0, 0, 800, 64, 1, 64
    
    If Not inChat Then
        If Not inTutorial Then
            ' render chatbox
            Width = 412
            height = 145
            'EngineRenderRectangle Tex_GUI(1), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_GUI(1), GUIWindow(GUI_CHAT).x, GUIWindow(GUI_CHAT).y, 0, 0, Width, height, Width, height
            RenderChatTextBuffer
            ' render the message input
            If chatOn Then RenderText Font_Default, RenderChatText & chatShowLine, GUIWindow(GUI_CHAT).x + 38, GUIWindow(GUI_CHAT).y + 126, White
            ' draw buttons
            For i = 34 To 35
                ' set co-ordinate
                x = GUIWindow(GUI_CHAT).x + Buttons(i).x
                y = GUIWindow(GUI_CHAT).y + Buttons(i).y
                Width = Buttons(i).Width
                height = Buttons(i).height
                ' check for state
                If Buttons(i).state = 2 Then
                    ' we're clicked boyo
                    'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                    RenderTexture Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
                ElseIf (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
                    ' we're hoverin'
                    'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                    RenderTexture Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
                    ' play sound if needed
                    If Not lastButtonSound = i Then
                        Play_Sound Sound_ButtonHover
                        lastButtonSound = i
                    End If
                Else
                    ' we're normal
                    'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                    RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
                    ' reset sound if needed
                    If lastButtonSound = i Then lastButtonSound = 0
                End If
            Next
        Else
            DrawTutorial
            Exit Sub
        End If
    Else
        DrawNpcChat
    End If
    
    ' render bars
    DrawGUIBars
    
    ' render menu
    DrawMenu
    
    ' render hotbar
    DrawHotbar
    
    ' render menus
    If GUIWindow(GUI_INVENTORY).visible Then DrawInventory
    If GUIWindow(GUI_SPELLS).visible Then DrawSkills
    If GUIWindow(GUI_CHARACTER).visible Then DrawCharacter
    If GUIWindow(GUI_OPTIONS).visible Then DrawOptions
    If GUIWindow(GUI_PARTY).visible Then DrawParty
    If GUIWindow(GUI_SHOP).visible Then DrawShop
    
    ' Drag and drop
    DrawDragItem
    DrawDragSpell
    
    ' Descriptions
    DrawInventoryItemDesc
    DrawCharacterItemDesc
    DrawPlayerSpellDesc
End Sub

Public Sub DrawTutorial()
Dim x As Long, y As Long, i As Long, Width As Long
Dim height As Long

    x = GUIWindow(GUI_CHAT).x
    y = GUIWindow(GUI_CHAT).y - 107
    
    ' render chatbox
    Width = 481
    height = 252
    'EngineRenderRectangle Tex_GUI(30), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(30), x, y, 0, 0, Width, height, Width, height
    
    ' Draw the text
    RenderText Font_Default, WordWrap(chatText, 260), x + 200, y + 129, White
    
    ' Draw replies
    For i = 1 To 4
        If Len(Trim$(chatOpt(i))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(i)) & "]")
            x = GUIWindow(GUI_CHAT).x + 200 + (130 - (Width / 2))
            y = GUIWindow(GUI_CHAT).y + 115 - ((i - 1) * 15)
            If chatOptState(i) = 2 Then
                ' clicked
                RenderText Font_Default, "[" & Trim$(chatOpt(i)) & "]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Default, "[" & Trim$(chatOpt(i)) & "]", x, y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = i Then
                        Play_Sound Sound_ButtonHover
                        lastNpcChatsound = i
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[" & Trim$(chatOpt(i)) & "]", x, y, BrightBlue
                    ' reset sound if needed
                    If lastNpcChatsound = i Then lastNpcChatsound = 0
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawNpcChat()
Dim i As Long, x As Long, y As Long, sprite As Long, Width As Long
Dim height As Long

    ' draw background
    x = GUIWindow(GUI_CHAT).x
    y = GUIWindow(GUI_CHAT).y
    
    ' render chatbox
    Width = 412
    height = 145
    'EngineRenderRectangle Tex_GUI(27), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(27), x, y, 0, 0, Width, height, Width, height
    
    ' draw the face
    sprite = Npc(chatNpc).sprite
    If sprite > 0 And sprite <= Count_Face Then
        'EngineRenderRectangle Tex_Face(sprite), x + 22, y + 22, 0, 0, 96, 96, 96, 96, 96, 96
        RenderTexture Tex_Face(sprite), x + 22, y + 22, 0, 0, 96, 96, 96, 96
    End If
    
    ' Draw the text
    RenderText Font_Default, WordWrap(chatText, 260), x + 130, y + 22, White
    
    ' Draw replies
    For i = 1 To 4
        If Len(Trim$(chatOpt(i))) > 0 Then
            Width = EngineGetTextWidth(Font_Default, "[" & Trim$(chatOpt(i)) & "]")
            x = GUIWindow(GUI_CHAT).x + 95 + (155 - (Width / 2))
            y = GUIWindow(GUI_CHAT).y + 115 - ((i - 1) * 15)
            If chatOptState(i) = 2 Then
                ' clicked
                RenderText Font_Default, "[" & Trim$(chatOpt(i)) & "]", x, y, Grey
            Else
                If (GlobalX >= x And GlobalX <= x + Width) And (GlobalY >= y And GlobalY <= y + 14) Then
                    ' hover
                    RenderText Font_Default, "[" & Trim$(chatOpt(i)) & "]", x, y, Yellow
                    ' play sound if needed
                    If Not lastNpcChatsound = i Then
                        Play_Sound Sound_ButtonHover
                        lastNpcChatsound = i
                    End If
                Else
                    ' normal
                    RenderText Font_Default, "[" & Trim$(chatOpt(i)) & "]", x, y, BrightBlue
                    ' reset sound if needed
                    If lastNpcChatsound = i Then lastNpcChatsound = 0
                End If
            End If
        End If
    Next
End Sub

Public Sub DrawShop()
Dim i As Long, x As Long, y As Long, itemnum As Long, itempic As Long, left As Long, top As Long, Amount As Long, colour As Long
Dim Width As Long, height As Long

    ' render the window
    Width = 252
    height = 317
    'EngineRenderRectangle Tex_GUI(28), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(28), GUIWindow(GUI_SHOP).x, GUIWindow(GUI_SHOP).y, 0, 0, Width, height, Width, height
    
    ' render the shop items
    For i = 1 To MAX_TRADES
        itemnum = Shop(InShop).TradeItem(i).Item
        If itemnum > 0 And itemnum <= MAX_ITEMS Then
            itempic = Item(itemnum).Pic
            If itempic > 0 And itempic <= Count_Item Then
                
                top = GUIWindow(GUI_SHOP).y + ShopTop + ((ShopOffsetY + 32) * ((i - 1) \ ShopColumns))
                left = GUIWindow(GUI_SHOP).x + ShopLeft + ((ShopOffsetX + 32) * (((i - 1) Mod ShopColumns)))
                
                'EngineRenderRectangle Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32, 32, 32
                RenderTexture Tex_Item(itempic), left, top, 0, 0, 32, 32, 32, 32
                
                ' If item is a stack - draw the amount you have
                If Shop(InShop).TradeItem(i).ItemValue > 1 Then
                    y = GUIWindow(GUI_SHOP).y + top + 22
                    x = GUIWindow(GUI_SHOP).x + left - 4
                    Amount = CStr(Shop(InShop).TradeItem(i).ItemValue)
                    
                    ' Draw currency but with k, m, b etc. using a convertion function
                    If CLng(Amount) < 1000000 Then
                        colour = White
                    ElseIf CLng(Amount) > 1000000 And CLng(Amount) < 10000000 Then
                        colour = Yellow
                    ElseIf CLng(Amount) > 10000000 Then
                        colour = BrightGreen
                    End If
                    
                    RenderText Font_Default, ConvertCurrency(Amount), x, y, colour
                End If
            End If
        End If
    Next
    
    ' draw buttons
    For i = 23 To 23
        ' set co-ordinate
        x = GUIWindow(GUI_SHOP).x + Buttons(i).x
        y = GUIWindow(GUI_SHOP).y + Buttons(i).y
        Width = Buttons(i).Width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ' play sound if needed
            If Not lastButtonSound = i Then
                Play_Sound Sound_ButtonHover
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
    
    ' draw item descriptions
    DrawShopItemDesc
End Sub

Public Sub DrawMenu()
Dim i As Long, x As Long, y As Long
Dim Width As Long, height As Long

    ' draw background
    x = GUIWindow(GUI_MENU).x
    y = GUIWindow(GUI_MENU).y
    Width = 232
    height = 76
    'EngineRenderRectangle Tex_GUI(3), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(3), x, y, 0, 0, Width, height, Width, height
    
    ' draw buttons
    For i = 1 To 6
        ' set co-ordinate
        x = GUIWindow(GUI_MENU).x + Buttons(i).x
        y = GUIWindow(GUI_MENU).y + Buttons(i).y
        Width = Buttons(i).Width
        height = Buttons(i).height
        ' check for state
        If Buttons(i).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ' play sound if needed
            If Not lastButtonSound = i Then
                Play_Sound Sound_ButtonHover
                lastButtonSound = i
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ' reset sound if needed
            If lastButtonSound = i Then lastButtonSound = 0
        End If
    Next
End Sub

Public Sub DrawMainMenu()
Dim i As Long, x As Long, y As Long
Dim Width As Long, height As Long
    
    ' draw logo
    Width = 503
    height = 172
    'EngineRenderRectangle Tex_GUI(36), 152, 20, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(36), 152, 20, 0, 0, Width, height, Width, height

    ' draw background
    x = GUIWindow(GUI_MAINMENU).x
    y = GUIWindow(GUI_MAINMENU).y
    Width = 495
    height = 332
    'EngineRenderRectangle Tex_GUI(18), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(18), x, y, 0, 0, Width, height, Width, height
    
    ' draw buttons
    If Not faderAlpha > 0 Then
        For i = 7 To 10
            ' set co-ordinate
            x = GUIWindow(GUI_MAINMENU).x + Buttons(i).x
            y = GUIWindow(GUI_MAINMENU).y + Buttons(i).y
            Width = Buttons(i).Width
            height = Buttons(i).height
            ' check for state
            If Buttons(i).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
            ElseIf (GlobalX >= x And GlobalX <= x + Buttons(i).Width) And (GlobalY >= y And GlobalY <= y + Buttons(i).height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
                ' play sound if needed
                If Not lastButtonSound = i Then
                    Play_Sound Sound_ButtonHover
                    lastButtonSound = i
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(i).PicNum), x, y, 0, 0, Width, height, Width, height
                ' reset sound if needed
                If lastButtonSound = i Then lastButtonSound = 0
            End If
        Next
    End If
    
    ' draw specific menus
    Select Case curMenu
        Case MENU_MAIN
            DrawNews
        Case MENU_LOGIN
            DrawLogin
        Case MENU_REGISTER
            DrawRegister
        Case MENU_CREDITS
            DrawCredits
        Case MENU_CLASS
            DrawClassSelect
        Case MENU_NEWCHAR
            DrawNewChar
    End Select
End Sub

Public Sub DrawNewChar()
Dim x As Long, y As Long, buttonnum As Long, sprite As Long
Dim Width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x
    y = GUIWindow(GUI_MAINMENU).y
    
    ' draw the image
    Width = 291
    height = 107
    'EngineRenderRectangle Tex_GUI(26), x + 110, y + 92, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(26), x + 110, y + 92, 0, 0, Width, height, Width, height
    
    ' char name
    RenderText Font_Default, sChar & chatShowLine, x + 158, y + 94, White
    
    ' sprite preview
    sprite = Class(newCharClass).MaleSprite(newCharSprite)
    'EngineRenderRectangle Tex_Char(sprite), x + 235, y + 123, 32, 0, 32, 32, 32, 32, 32, 32
    RenderTexture Tex_Char(sprite), x + 235, y + 123, 32, 0, 32, 32, 32, 32
    
    If Not faderAlpha > 0 Then
        ' position
        buttonnum = 15
        x = GUIWindow(GUI_MAINMENU).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_MAINMENU).y + Buttons(buttonnum).y
        Width = Buttons(buttonnum).Width
        height = Buttons(buttonnum).height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                Play_Sound Sound_ButtonHover
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    End If
End Sub

Public Sub DrawClassSelect()
Dim x As Long, y As Long, buttonnum As Long
Dim Width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x
    y = GUIWindow(GUI_MAINMENU).y
    
    Select Case newCharClass
        Case 1 ' warrior
            Width = 426
            height = 209
            'EngineRenderRectangle Tex_GUI(23), x + 30, y + 34, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_GUI(23), x + 30, y + 34, 0, 0, Width, height, Width, height
        Case 2 ' wizard
            Width = 441
            height = 213
            'EngineRenderRectangle Tex_GUI(24), x + 30, y + 33, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_GUI(24), x + 30, y + 33, 0, 0, Width, height, Width, height
        Case 3 ' whisperer
            Width = 455
            height = 212
            'EngineRenderRectangle Tex_GUI(25), x + 30, y + 38, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_GUI(25), x + 30, y + 38, 0, 0, Width, height, Width, height
    End Select
    
    If Not faderAlpha > 0 Then
        For buttonnum = 13 To 14
            x = GUIWindow(GUI_MAINMENU).x + Buttons(buttonnum).x
            y = GUIWindow(GUI_MAINMENU).y + Buttons(buttonnum).y
            Width = Buttons(buttonnum).Width
            height = Buttons(buttonnum).height
            ' render accept button
            If Buttons(buttonnum).state = 2 Then
                ' we're clicked boyo
                'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
            ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).height) Then
                ' we're hoverin'
                'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
                ' play sound if needed
                If Not lastButtonSound = buttonnum Then
                    Play_Sound Sound_ButtonHover
                    lastButtonSound = buttonnum
                End If
            Else
                ' we're normal
                'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
                RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
                ' reset sound if needed
                If lastButtonSound = buttonnum Then lastButtonSound = 0
            End If
        Next
    End If
End Sub

Public Sub DrawNews()
Dim x As Long, y As Long
Dim Width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x + 137
    y = GUIWindow(GUI_MAINMENU).y + 80
    Width = 224
    height = 118
    'EngineRenderRectangle Tex_GUI(22), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(22), x, y, 0, 0, Width, height, Width, height
End Sub

Public Sub DrawLogin()
Dim x As Long, y As Long, buttonnum As Long
Dim Width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x + 86
    y = GUIWindow(GUI_MAINMENU).y + 102
    buttonnum = 11
    
    ' render block
    Width = 317
    height = 94
    'EngineRenderRectangle Tex_GUI(21), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(21), x, y, 0, 0, Width, height, Width, height
    
    ' render username
    If curTextbox = 1 Then ' focuses
        RenderText Font_Default, sUser & chatShowLine, x + 74, y + 2, White
    Else
        RenderText Font_Default, sUser, x + 74, y + 2, White
    End If
    
    ' render password
    If curTextbox = 2 Then ' focuses
        RenderText Font_Default, CensorWord(sPass) & chatShowLine, x + 74, y + 26, White
    Else
        RenderText Font_Default, CensorWord(sPass), x + 74, y + 26, White
    End If
    
    If Not faderAlpha > 0 Then
        ' position
        x = GUIWindow(GUI_MAINMENU).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_MAINMENU).y + Buttons(buttonnum).y
        Width = Buttons(buttonnum).Width
        height = Buttons(buttonnum).height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(buttonnum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(buttonnum), x, y, 0, 0, Width, height, Width, height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                Play_Sound Sound_ButtonHover
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    End If
End Sub

Public Sub DrawRegister()
Dim x As Long, y As Long, buttonnum As Long
Dim Width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x + 86
    y = GUIWindow(GUI_MAINMENU).y + 92
    buttonnum = 12
    
    ' render block
    Width = 319
    height = 107
    'EngineRenderRectangle Tex_GUI(20), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(20), x, y, 0, 0, Width, height, Width, height
    
    ' render username
    If curTextbox = 1 Then ' focuses
        RenderText Font_Default, sUser & chatShowLine, x + 74, y + 2, White
    Else
        RenderText Font_Default, sUser, x + 74, y + 2, White
    End If
    
    ' render password
    If curTextbox = 2 Then ' focuses
        RenderText Font_Default, CensorWord(sPass) & chatShowLine, x + 74, y + 26, White
    Else
        RenderText Font_Default, CensorWord(sPass), x + 74, y + 26, White
    End If
    
    ' render password
    If curTextbox = 3 Then ' focuses
        RenderText Font_Default, CensorWord(sPass2) & chatShowLine, x + 74, y + 50, White
    Else
        RenderText Font_Default, CensorWord(sPass2), x + 74, y + 50, White
    End If
    
    If Not faderAlpha > 0 Then
        ' position
        x = GUIWindow(GUI_MAINMENU).x + Buttons(buttonnum).x
        y = GUIWindow(GUI_MAINMENU).y + Buttons(buttonnum).y
        Width = Buttons(buttonnum).Width
        height = Buttons(buttonnum).height
        ' render accept button
        If Buttons(buttonnum).state = 2 Then
            ' we're clicked boyo
            'EngineRenderRectangle Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_c(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
        ElseIf (GlobalX >= x And GlobalX <= x + Buttons(buttonnum).Width) And (GlobalY >= y And GlobalY <= y + Buttons(buttonnum).height) Then
            ' we're hoverin'
            'EngineRenderRectangle Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons_h(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
            ' play sound if needed
            If Not lastButtonSound = buttonnum Then
                Play_Sound Sound_ButtonHover
                lastButtonSound = buttonnum
            End If
        Else
            ' we're normal
            'EngineRenderRectangle Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, width, height, width, height, width, height
            RenderTexture Tex_Buttons(Buttons(buttonnum).PicNum), x, y, 0, 0, Width, height, Width, height
            ' reset sound if needed
            If lastButtonSound = buttonnum Then lastButtonSound = 0
        End If
    End If
End Sub

Public Sub DrawCredits()
Dim x As Long, y As Long
Dim Width As Long, height As Long

    x = GUIWindow(GUI_MAINMENU).x + 187
    y = GUIWindow(GUI_MAINMENU).y + 86
    Width = 121
    height = 104
    'engineRenderRectangle Tex_GUI(19), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(19), x, y, 0, 0, Width, height, Width, height
End Sub

Public Sub DrawGUIBars()
Dim tmpWidth As Long, barWidth As Long, x As Long, y As Long, dX As Long, dY As Long, sString As String
Dim Width As Long, height As Long

    ' backwindow + empty bars
    x = GUIWindow(GUI_BARS).x
    y = GUIWindow(GUI_BARS).y
    Width = 254
    height = 75
    'EngineRenderRectangle Tex_GUI(4), x, y, 0, 0, width, height, width, height, width, height
    RenderTexture Tex_GUI(4), x, y, 0, 0, Width, height, Width, height
    
    ' hardcoded for POT textures
    barWidth = 241
    
    ' health bar
    BarWidth_GuiHP_Max = ((GetPlayerVital(MyIndex, Vitals.HP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / barWidth)) * barWidth
    RenderTexture Tex_GUI(13), x + 7, y + 9, 0, 0, BarWidth_GuiHP, D3DT_TEXTURE(Tex_GUI(13)).height, BarWidth_GuiHP, D3DT_TEXTURE(Tex_GUI(13)).height
    ' render health
    sString = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
    dX = x + 7 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = y + 9
    RenderText Font_Default, sString, dX, dY, White
    
    ' spirit bar
    BarWidth_GuiSP_Max = ((GetPlayerVital(MyIndex, Vitals.MP) / barWidth) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / barWidth)) * barWidth
    RenderTexture Tex_GUI(14), x + 7, y + 31, 0, 0, BarWidth_GuiSP, D3DT_TEXTURE(Tex_GUI(14)).height, BarWidth_GuiSP, D3DT_TEXTURE(Tex_GUI(14)).height
    ' render spirit
    sString = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
    dX = x + 7 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = y + 31
    RenderText Font_Default, sString, dX, dY, White
    
    ' exp bar
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        BarWidth_GuiEXP_Max = ((GetPlayerExp(MyIndex) / barWidth) / (TNL / barWidth)) * barWidth
    Else
        BarWidth_GuiEXP_Max = barWidth
    End If
    RenderTexture Tex_GUI(15), x + 7, y + 53, 0, 0, BarWidth_GuiEXP, D3DT_TEXTURE(Tex_GUI(15)).height, BarWidth_GuiEXP, D3DT_TEXTURE(Tex_GUI(15)).height
    ' render exp
    If GetPlayerLevel(MyIndex) < MAX_LEVELS Then
        sString = GetPlayerExp(MyIndex) & "/" & TNL
    Else
        sString = "Max Level"
    End If
    dX = x + 7 + (barWidth / 2) - (EngineGetTextWidth(Font_Default, sString) / 2)
    dY = y + 53
    RenderText Font_Default, sString, dX, dY, White
End Sub

Public Sub DrawGDI()
    If frmEditor_Animation.visible Then
        GDIRenderAnimation
    ElseIf frmEditor_Item.visible Then
        GDIRenderItem frmEditor_Item.picItem, frmEditor_Item.scrlPic.value
    ElseIf frmEditor_Map.visible Then
        GDIRenderTileset
    ElseIf frmEditor_NPC.visible Then
        GDIRenderChar frmEditor_NPC.picSprite, frmEditor_NPC.scrlSprite.value
    ElseIf frmEditor_Resource.visible Then
        ' lol nothing
    ElseIf frmEditor_Spell.visible Then
        GDIRenderSpell frmEditor_Spell.picSprite, frmEditor_Spell.scrlIcon.value
    End If
End Sub

' Main Loop
Public Sub Render_Graphics()
Dim x As Long, y As Long, i As Long

    ' fuck off if we're not doing anything
    If GettingMap Then Exit Sub

    ' update the camera
    UpdateCamera
    
    ' make sure we've got control of the form
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then
        If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Then
            Exit Sub
        End If
        Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
        Call D3DDevice8.Reset(D3DWindow)
        Call EngineInitRenderStates
    End If
    
    ' unload any textures we need to unload
    UnloadTextures
    
    ' Start rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    ' render lower tiles
    If Count_Tileset > 0 Then
        For x = TileView.left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapTile(x, y)
                End If
            Next
        Next
    End If
    
    ' render the items
    If Count_Item > 0 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                Call DrawItem(i)
            End If
        Next
    End If
    
    ' draw animations
    If Count_Anim > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(0) Then
                DrawAnimation i, 0
            End If
        Next
    End If
    
    ' Y-based render. Renders Players, Npcs and Resources based on Y-axis.
    If Count_Char > 0 Then
        ' shadows - Players
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call DrawShadow(Player(i).sprite, (Player(i).x * 32) + Player(i).xOffset, (Player(i).y * 32) + Player(i).yOffset)
            End If
        Next
        
        ' shadows - npcs
        For i = 1 To MAX_MAP_NPCS
            If MapNpc(i).num > 0 Then
                Call DrawShadow(Npc(MapNpc(i).num).sprite, (MapNpc(i).x * 32) + MapNpc(i).xOffset, (MapNpc(i).y * 32) + MapNpc(i).yOffset)
            End If
        Next
        
        ' Players
        For i = 1 To MAX_PLAYERS
            If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
                Call DrawPlayer(i)
            End If
        Next
        
        ' Npcs
        For i = 1 To MAX_MAP_NPCS
            Call DrawNpc(i)
        Next
    End If
    
    ' Resources
    If Count_Resource > 0 Then
        If Resources_Init Then
            If Resource_Index > 0 Then
                For i = 1 To Resource_Index
                    Call DrawResource(i)
                Next
            End If
        End If
    End If
    
    ' render out upper tiles
    If Count_Tileset > 0 Then
        For x = TileView.left To TileView.Right
            For y = TileView.top To TileView.bottom
                If IsValidMapPoint(x, y) Then
                    Call DrawMapFringeTile(x, y)
                End If
            Next
        Next
    End If
    
    ' render fog
    DrawFog
    
    ' render animations
    If Count_Anim > 0 Then
        For i = 1 To MAX_BYTE
            If AnimInstance(i).Used(1) Then
                DrawAnimation i, 1
            End If
        Next
    End If
    
    ' render target
    If myTarget > 0 Then
        If myTargetType = TARGET_TYPE_PLAYER Then
            DrawTarget (Player(myTarget).x * 32) + Player(myTarget).xOffset, (Player(myTarget).y * 32) + Player(myTarget).yOffset
        ElseIf myTargetType = TARGET_TYPE_NPC Then
            DrawTarget (MapNpc(myTarget).x * 32) + MapNpc(myTarget).xOffset, (MapNpc(myTarget).y * 32) + MapNpc(myTarget).yOffset
        End If
    End If
    
    ' blt the hover icon
    DrawTargetHover
    
    ' draw the bars
    DrawBars
    
    ' draw attributes
    If InMapEditor Then
        DrawMapAttributes
    End If
    
    ' draw player names
    For i = 1 To MAX_PLAYERS
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            Call DrawPlayerName(i)
        End If
    Next
    
    ' draw npc names
    For i = 1 To MAX_MAP_NPCS
        If MapNpc(i).num > 0 Then
            Call DrawNpcName(i)
        End If
    Next
    
    ' draw action msg
    For i = 1 To MAX_BYTE
        DrawActionMsg i
    Next
    
    If InMapEditor Then
        If frmEditor_Map.optBlock.value = True Then
            For x = TileView.left To TileView.Right
                For y = TileView.top To TileView.bottom
                    If IsValidMapPoint(x, y) Then
                        Call DrawDirection(x, y)
                    End If
                Next
            Next
        End If
    End If
    
    ' draw the messages
    For i = 1 To MAX_BYTE
        If chatBubble(i).active Then
            DrawChatBubble i
        End If
    Next
    
    ' Draw the GUI
    If Not InMapEditor And Not hideGUI Then DrawGUI
    
    ' Draw fade in
    If canFade Then DrawFader
    
    ' render fps
    If BFPS Then
        RenderText Font_Default, "FPS: " & GameFPS, 0, 0, White
        RenderText Font_Default, "Latency: " & Ping / 2, 0, 15, White
    End If
    
    ' draw loc
    If BLoc Then
        RenderText Font_Default, Trim$("cur x: " & CurX & " y: " & CurY), Camera.left, Camera.top, Yellow
        RenderText Font_Default, Trim$("loc x: " & GetPlayerX(MyIndex) & " y: " & GetPlayerY(MyIndex)), Camera.left, Camera.top + 16, Yellow
        RenderText Font_Default, Trim$(" (map #" & GetPlayerMap(MyIndex) & ")"), Camera.left, Camera.top + 32, Yellow
    End If
    
    ' End the rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
    
    ' GDI Rendering
    DrawGDI
End Sub

Public Sub DrawFader()
    If faderAlpha < 0 Then faderAlpha = 0
    If faderAlpha > 254 Then faderAlpha = 254
    'EngineRenderRectangle 0, 0, 0, 0, 0, 800, 600, 0, 0, 800, 600, 0, 0, 0, 0, , , faderAlpha, 0, 0, 0
    RenderTexture Tex_Fader, 0, 0, 0, 0, 800, 600, 32, 32, D3DColorARGB(faderAlpha, 0, 0, 0)
End Sub

Public Sub Render_Menu()
    ' make sure we've got control of the form
    If D3DDevice8.TestCooperativeLevel <> D3D_OK Then
        If D3DDevice8.TestCooperativeLevel = D3DERR_DEVICELOST Then
            Exit Sub
        End If
        Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
        Call D3DDevice8.Reset(D3DWindow)
        Call EngineInitRenderStates
    End If
    
    ' unload any textures we need to unload
    UnloadTextures
    
    ' Start rendering
    Call D3DDevice8.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice8.BeginScene
    
    ' fader
    Select Case faderState
        Case 0, 1
            ' render background
            If Not faderAlpha = 255 Then EngineRenderSurface 1, 0, 0, 0, 0, 800, 600
            ' fading in/out to first screen
            DrawFader
        Case 2, 3
            ' render background
            If Not faderAlpha = 255 Then EngineRenderSurface 2, 0, 0, 0, 0, 800, 600
            ' fading in to second screen
            DrawFader
    End Select
    
    ' render menu
    If faderState >= 4 And Not faderAlpha = 255 Then
        ' render background
        EngineRenderSurface 3, 0, 0, 0, 0, 800, 600
        
        ' render menu block
        DrawMainMenu
    End If
    
    ' render last fader
    If faderState >= 4 Then
        ' fading in to menu
        If Not faderAlpha = 255 Then DrawFader
    End If
    
    ' End the rendering
    Call D3DDevice8.EndScene
    Call D3DDevice8.Present(ByVal 0, ByVal 0, 0, ByVal 0)
End Sub
