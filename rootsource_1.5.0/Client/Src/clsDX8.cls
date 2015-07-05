VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDX8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private DX8 As DirectX8
Private D3D As Direct3D8
Private D3DX As D3DX8
Private D3DDevice As Direct3DDevice8

Private Const FVF As Long = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

Private D3DT_TEXTURE() As Direct3DTexture8

Private Const FVF_ShaderVertex As Long = D3DFVF_XYZRHW Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR Or D3DFVF_TEX2
Public DX_VB As Direct3DVertexBuffer8

Private Const TEXTURE_NULL As Long = 0

Private CurrentTexture As Long
Private TEXTURE_FONT As Long
Private mTextureNum As Long

Private Type rectangle
    Top As Long
    Bottom As Long
    Left As Long
    Right As Long
End Type

Private MainFont As D3DXFont '//This will be *created* in a minute
Private MainFontDesc As IFont '//We use this temporarily to setup the font
Private TextRect As RECT '//This defines where it will be
Public fnt As New StdFont '//This is also used to describe and setup the font


Private Type VERTEXx
    x As Single
    y As Single
    Z As Single
    RHW As Single
    Color As Long
    Specular As Long
    TU As Single
    TV As Single
End Type

Private Box(0 To 3) As VERTEX

Private Const Pi As Single = 3.141592653589
Private Const Deg2Rad As Single = Pi / 180#

Public Function InitDirectX(ByVal hWnd As Long, BehaviorFlags As CONST_D3DCREATEFLAGS) As Long
Dim DispMode As D3DDISPLAYMODE
Dim D3DWindow As D3DPRESENT_PARAMETERS

On Error GoTo ErrorInit

    Set DX8 = New DirectX8
    Set D3D = DX8.Direct3DCreate()
    
    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    DispMode.Format = D3DFMT_A8R8G8B8
    
    D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY
    D3DWindow.hDeviceWindow = hWnd
    D3DWindow.BackBufferFormat = DispMode.Format
    D3DWindow.Windowed = 1
    
    Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, BehaviorFlags, D3DWindow)
    Select Case BehaviorFlags
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING
            Debug.Print "D3D created with D3DCREATE_HARDWARE_VERTEXPROCESSING"
        Case D3DCREATE_MIXED_VERTEXPROCESSING
            Debug.Print "D3D created with D3DCREATE_MIXED_VERTEXPROCESSING"
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            Debug.Print "D3D created with D3DCREATE_SOFTWARE_VERTEXPROCESSING"
    End Select
    
    Set D3DX = New D3DX8
    
    Call D3DDevice.SetVertexShader(FVF)
    
    ReDim D3DT_TEXTURE(0 To 0) As Direct3DTexture8
    
    TEXTURE_FONT = LoadTexture(App.Path & "\gfx\Courier12Bold.png", 512, 512)
    
    Dim i As Long
    
    Do While FileExist("gfx\tiles" & i & ".png")
        
        ReDim Preserve Tr_Tiles(0 To i)
        Tr_Tiles(i) = LoadTexture(App.Path & "\gfx\tiles" & i & ".png", 256, 256)
        
        i = i + 1
    Loop
    
    TileCount = i - 1
    
    i = 1
    
    Do While FileExist("gfx\items\" & i & ".png")
        
        ReDim Preserve Tr_Items(1 To i)
        Tr_Items(i) = LoadTexture(App.Path & "\gfx\items\" & i & ".png", 32, 32)
        
        i = i + 1
    Loop
    
    ItemCount = i - 1
    
    i = 1
    
    Do While FileExist("gfx\sprites\" & i & ".png")
        
        ReDim Preserve Tr_Sprites(1 To i)
        Tr_Sprites(i) = LoadTexture(App.Path & "\gfx\sprites\" & i & ".png", 512, 32)
        
        i = i + 1
    Loop
    
    SpriteCount = i - 1
    
    i = 1
    
    Do While FileExist("gfx\spells\" & i & ".png")
        
        ReDim Preserve Tr_Spells(1 To i)
        Tr_Spells(i) = LoadTexture(App.Path & "\gfx\spells\" & i & ".png", 512, 32)
        
        i = i + 1
    Loop
    
    SpellCount = i - 1
    
    fnt.Name = "Courier"
    fnt.size = 12
    fnt.Bold = True
    Set MainFontDesc = fnt

    Set MainFont = D3DX.CreateFont(D3DDevice, MainFontDesc.hFont)
    
    Call D3DDevice.SetRenderState(D3DRS_ALPHABLENDENABLE, True)
    Call D3DDevice.SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
    Call D3DDevice.SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_MODULATE)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
    Call D3DDevice.SetTextureStageState(0, D3DTSS_ALPHAARG1, 2)
    
    InitDirectX = 1
    
    Exit Function

ErrorInit:
    Select Case BehaviorFlags
        Case D3DCREATE_HARDWARE_VERTEXPROCESSING
            Debug.Print "D3DCREATE_HARDWARE_VERTEXPROCESSING not supported"
        Case D3DCREATE_MIXED_VERTEXPROCESSING
            Debug.Print "D3DCREATE_MIXED_VERTEXPROCESSING not supported"
        Case D3DCREATE_SOFTWARE_VERTEXPROCESSING
            Debug.Print "D3DCREATE_SOFTWARE_VERTEXPROCESSING not supported"
    End Select
    InitDirectX = 0
    Exit Function
End Function

Public Sub BeginScene()
    Call D3DDevice.Clear(0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0)
    Call D3DDevice.BeginScene
End Sub

Public Sub EndScene()
    Call D3DDevice.EndScene
    Call D3DDevice.Present(ByVal 0, ByVal 0, 0, ByVal 0)
End Sub

Public Sub EndSceneSp(ByVal hWnd As Long, ByRef SrcRect As D3DRECT)
    Call D3DDevice.EndScene
    Call D3DDevice.Present(SrcRect, ByVal 0, hWnd, ByVal 0)
End Sub

Public Sub EndSceneNP()
    Call D3DDevice.EndScene
End Sub

Public Function LoadTexture(ByVal Path As String, ByVal width As Long, ByVal height As Long) As Long
    mTextureNum = mTextureNum + 1
    ReDim Preserve D3DT_TEXTURE(0 To mTextureNum) As Direct3DTexture8
    Set D3DT_TEXTURE(mTextureNum) = D3DX.CreateTextureFromFileEx(D3DDevice, Path, width, height, D3DX_DEFAULT, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, 0, ByVal 0, ByVal 0)
    LoadTexture = mTextureNum
End Function

Public Sub SetTexture(ByVal Texture As Long)
    If Texture <> CurrentTexture Then
        If Texture > UBound(D3DT_TEXTURE) Then Texture = UBound(D3DT_TEXTURE)
        If Texture < 0 Then Texture = 0
        Call D3DDevice.SetTexture(0, D3DT_TEXTURE(Texture))
        CurrentTexture = Texture
    End If
End Sub

Public Sub GeomSet(ByVal x As Single, ByVal y As Single, ByVal w As Integer, ByVal h As Integer, ByVal SX As Single, ByVal SY As Single, ByVal SW As Single, ByVal SH As Single, ByVal colour As Long, WM As Long, HM As Long)
Dim TempX As Single, TempY As Single
Dim SinQ As Single, CosQ As Single
Dim i As Long

    SH = 1 / HM * (SH + SY)
    SW = 1 / WM * (SW + SX)
    SX = 1 / WM * SX
    SY = 1 / HM * SY
    
    Box(0) = MakeVertex(x, y, 0, 1, colour, 0, SX, SY)
    Box(1) = MakeVertex(x + w, y, 0, 1, colour, 0, SW, SY)
    Box(2) = MakeVertex(x, y + h, 0, 1, colour, 0, SX, SH)
    Box(3) = MakeVertex(x + w, y + h, 0, 1, colour, 0, SW, SH)
End Sub

Public Sub GeomRotate(ByVal x As Single, ByVal y As Single, ByVal w As Integer, ByVal h As Integer, Optional ByVal Angle As Long, Optional ByVal OW As Long, Optional ByVal OH As Long)
Dim TempX As Single, TempY As Single
Dim SinQ As Single, CosQ As Single
Dim i As Long

    If Angle <> 0 Then
        SinQ = Sin(Angle * Deg2Rad)
        CosQ = Cos(Angle * Deg2Rad)
        
        For i = 0 To 3
            TempX = ((Box(i).x - x - OW) * CosQ) - ((Box(i).y - y - OH) * SinQ)
            TempY = ((Box(i).x - x - OW) * SinQ) + ((Box(i).y - y - OH) * CosQ)
            Box(i).x = TempX + x + OW
            Box(i).y = TempY + y + OH
        Next
    End If
End Sub

Public Sub GeomSetBox(ByVal x As Single, ByVal y As Single, ByVal w As Integer, ByVal h As Integer, ByVal colour As Long)
Dim TempX As Single, TempY As Single
Dim SinQ As Single, CosQ As Single
Dim i As Long

    Box(0) = MakeVertex(x, y, 0, 1, colour, 0, 0, 0)
    Box(1) = MakeVertex(x + w, y, 0, 1, colour, 0, 0, 0)
    Box(2) = MakeVertex(x, y + h, 0, 1, colour, 0, 0, 0)
    Box(3) = MakeVertex(x + w, y + h, 0, 1, colour, 0, 0, 0)
End Sub

Public Sub DrawTexture(ByVal x As Single, ByVal y As Single, ByVal w As Integer, ByVal h As Integer, ByVal SX As Single, ByVal SY As Single, ByVal SW As Single, ByVal SH As Single, ByVal colour As Long, WM As Long, HM As Long)
    Call GeomSet(x, y, w, h, SX, SY, SW, SH, colour, WM, HM)
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
End Sub

Public Sub DrawTextureNoGeom()
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
End Sub

Public Sub DrawBox(ByVal x As Single, ByVal y As Single, ByVal w As Integer, ByVal h As Integer, ByVal colour As Long, Optional ByVal Angle As Long = 0, Optional ByVal OW As Long = 16, Optional ByVal OH As Long = 16)
    Call SetTexture(TEXTURE_NULL)
    Call GeomSetBox(x, y, w, h, colour)
    If Angle <> 0 Then
        Call GeomRotate(x, y, w, h, Angle, OW, OH)
    End If
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
End Sub

Public Sub DrawBoxNoGeom()
    Call SetTexture(TEXTURE_NULL)
    Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
End Sub

Public Sub DrawText(ByVal x As Long, ByVal y As Long, ByRef oText As String, ByVal colour As Long)
Dim i As Long
Dim Text() As Byte

Text = ToArray(oText)

    Call SetTexture(TEXTURE_FONT)
    For i = 0 To UBound(Text)
        Call GeomSet(x + 1 + i * 10, y + 1, 12, 20, (Text(i) Mod 32) * 13, (Text(i) \ 32) * 17, 12, 20, &HFF000000, 512, 512)
        Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
        
        Call GeomSet(x + i * 10, y, 12, 20, (Text(i) Mod 32) * 13, (Text(i) \ 32) * 17, 12, 20, colour, 512, 512)
        Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
    Next
End Sub

Public Sub DrawTextx(ByVal x As Long, ByVal y As Long, ByVal Text As String, ByVal colour As Long, Optional shadow As Boolean = True, Optional Flags As Long = DT_SINGLELINE, Optional w As Long = 512, Optional h As Long = 384)
    If shadow Then
        TextRect.Top = y + 2
        TextRect.Left = x
        TextRect.Bottom = y + h
        TextRect.Right = x + w
        D3DX.DrawText MainFont, ARGB(255, 0, 0, 0), Text, TextRect, Flags
        TextRect.Top = y + 1
        TextRect.Left = x + 1
        TextRect.Bottom = y + h
        TextRect.Right = x + w
        D3DX.DrawText MainFont, ARGB(255, 0, 0, 0), Text, TextRect, Flags
    End If
    
    TextRect.Top = y
    TextRect.Left = x
    TextRect.Bottom = y + h
    TextRect.Right = x + w
    D3DX.DrawText MainFont, colour, Text, TextRect, Flags
End Sub

Public Sub DrawTextColourArray(ByVal x As Long, ByVal y As Long, ByRef Text() As Byte, ByRef colour() As Long)
Dim i As Long

    For i = 0 To UBound(Text)
        Call GeomSet(x + i * 6, y, 5, 8, Text(i) * 5, 0, 5, 8, colour(i), 512, 8)
        Call D3DDevice.DrawPrimitiveUP(D3DPT_TRIANGLESTRIP, 2, Box(0), Len(Box(0)))
    Next
End Sub

Private Function MakeVertex(x As Single, y As Single, Z As Single, RHW As Single, Color As Long, Specular As Long, TU As Single, TV As Single) As VERTEX
    MakeVertex.x = x
    MakeVertex.y = y
    MakeVertex.Z = Z
    MakeVertex.RHW = RHW
    MakeVertex.Color = Color
    MakeVertex.Specular = Specular
    MakeVertex.TU = TU
    MakeVertex.TV = TV
End Function

Public Function ARGB(ByVal a As Long, ByVal R As Long, ByVal G As Long, ByVal B As Long) As Long
    If a > 127 Then a = a Or -128
    ARGB = (a * &H1000000) + (R * &H10000) + (G * &H100) + B
End Function

Public Function AcRGB(ByVal a As Long, ByVal rgb As Long) As Long
    If a > 127 Then a = a Or -128
    AcRGB = (a * &H1000000) + rgb
End Function

Public Function ToArray(ByVal Text As String) As Byte()
Dim bArray() As Byte
Dim i As Long

    bArray = StrConv(Text, vbFromUnicode)
    For i = 0 To Len(Text) - 1
        If bArray(i) >= 32 Then
            bArray(i) = bArray(i) - 32
        Else
            bArray(i) = 0
        End If
    Next
    ToArray = bArray
End Function

Public Function ToArrayRange(ByRef bArray() As Long, ByVal Var As Long, ByVal Lower As Long, ByVal Upper As Long) As Long()
Dim i As Long

    For i = Lower To Upper
        bArray(i) = Var
    Next
    ToArrayRange = bArray
End Function

Public Function ToByteArray(ByVal Value As Byte) As Byte()
Dim bArray(0) As Byte

    bArray(0) = Value
    ToByteArray = bArray
End Function
