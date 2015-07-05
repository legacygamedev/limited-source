VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDirect3D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Enum Alignment
    DT_LEFT = 0
    DT_TOP = 1
    DT_CENTER = 2
    DT_RIGHT = 3
    DT_VCENTER = 4
    DT_BOTTOM = 5
    DT_SINGLELINE = 6
End Enum

Public d3d As Direct3D8
Public d3dx As New D3DX8
Private dispmode As D3DDISPLAYMODE
'public d3dpp As D3DPRESENT_PARAMETERS
Public d3ddev As Direct3DDevice8
Public d3dsprite As D3DXSprite
Public MainFont As D3DXFont '//This will be *created* in a minute
Public MainFontDesc As IFont '//We use this temporarily to setup the font
Public fnt As New StdFont

Private Const Transparency = &HFF00FF

Public Sub InitDirect3D(ByVal hwnd As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal bFullscreen As Boolean, ByVal bSetup As Boolean)
'On Error GoTo ErrorHandler:

'Set the Direct3D object
Set d3d = DX8.Direct3DCreate()

'--Set the display settings--
Dim d3dpp As D3DPRESENT_PARAMETERS
d3dpp.hDeviceWindow = hwnd 'Set the handle
d3dpp.BackBufferCount = 1 'Set the count
d3dpp.BackBufferWidth = lWidth 'Set the width
d3dpp.BackBufferHeight = lHeight 'Set the height
d3dpp.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
d3dpp.MultiSampleType = D3DMULTISAMPLE_NONE
d3dpp.AutoDepthStencilFormat = D3DFMT_D32

'Find the color depth according to the desktop
d3d.GetAdapterDisplayMode D3DADAPTER_DEFAULT, dispmode
d3dpp.BackBufferFormat = dispmode.Format

'Added just in case fullscreen ever used
If bFullscreen Then
    d3dpp.Windowed = 0
Else
    d3dpp.Windowed = 1
End If

'Create the D3D primary device
    Set d3ddev = d3d.CreateDevice( _
        D3DADAPTER_DEFAULT, _
        D3DDEVTYPE_HAL, _
        hwnd, _
        D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
        d3dpp)

'Initalize the sprite class
Set d3dsprite = d3dx.CreateSprite(d3ddev)

Call SetupFont("Verdana", 12, False, False)

If bSetup = True Then Call SetupGlobals

ErrorHandler:
'Debug.Print Err.Number
'Call GameDestroy
End Sub

Public Sub UnloadDirect3D(ByVal bSetup As Boolean)
If bSetup Then
    Set itemtex = Nothing
    Set spritetex = Nothing
    Set tiletex = Nothing
    Set mapeditortex = Nothing
End If
Set d3ddev = Nothing
Set d3d = Nothing
End Sub

Public Function LoadSurface(ByVal FileName As String, ByVal Width As Long, ByVal Height As Long, ByVal Filter As Long) As Direct3DSurface8
Dim surface As Direct3DSurface8

Set LoadSurface = Nothing

Set surface = d3ddev.CreateImageSurface(Width, Height, dispmode.Format)

d3dx.LoadSurfaceFromFile surface, ByVal 0, ByVal 0, FileName, ByVal 0, D3DX_DEFAULT, Filter, ByVal 0

Set LoadSurface = surface
End Function

Public Function LoadTexture(ByRef dev As Direct3DDevice8, ByVal FileName As String, ByVal Filter As Long) As Direct3DTexture8
Dim texture As Direct3DTexture8

Set texture = d3dx.CreateTextureFromFileEx(dev, FileName, D3DX_DEFAULT, D3DX_DEFAULT, 1, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_NONE, D3DX_FILTER_NONE, Filter, ByVal 0, ByVal 0)

Set LoadTexture = texture
End Function

Public Function GetPicWidth(psPath As String) As Long
  Dim f As Integer
  Dim tmp As String
  Dim FileHeader As BITMAPFILEHEADER
  Dim InfoHeader As BITMAPINFOHEADER
  
  On Error Resume Next
  GetPicWidth = 0
 
  f = FreeFile
  Open psPath For Binary Access Read As #f
  Get #f, , FileHeader
  Get #f, , InfoHeader
  Close #f
  
  GetPicWidth = InfoHeader.biWidth
End Function

Public Function GetPicHeight(psPath As String) As Long
  Dim f As Integer
  Dim tmp As String
  Dim FileHeader As BITMAPFILEHEADER
  Dim InfoHeader As BITMAPINFOHEADER
  
  On Error Resume Next
  GetPicHeight = 0
 
  f = FreeFile
  Open psPath For Binary Access Read As #f
  Get #f, , FileHeader
  Get #f, , InfoHeader
  Close #f
  
  GetPicHeight = InfoHeader.biHeight
End Function

Private Sub SetupGlobals()
Dim Path As String

'Begin loading the textures and surfaces
Path = App.Path & GFX_PATH & "Items.bmp"
    Set itemtex = LoadTexture(d3ddev, Path, Transparency)
Path = App.Path & GFX_PATH & "Sprites.bmp"
    Set spritetex = LoadTexture(d3ddev, Path, Transparency)
Path = App.Path & GFX_PATH & "Tiles.bmp"
    Set tiletex = LoadTexture(d3ddev, Path, Transparency)
Path = App.Path & GFX_PATH & "attribute_display.bmp"
    Set mapeditortex = LoadTexture(d3ddev, Path, Transparency)
' Define global colors
Call DefineColors
End Sub

Public Sub SetupFont(ByVal Name As String, ByVal Size As Byte, ByVal Bold As Boolean, ByVal Italic As Boolean)
'Initalize the font
fnt.Name = Name
fnt.Size = Size
If Bold = True Then
    fnt.Bold = True
Else
    fnt.Bold = False
End If
If Italic = True Then
    fnt.Italic = True
Else
    fnt.Italic = False
End If
If Not (MainFontDesc Is Nothing) Then Set MainFontDesc = Nothing
Set MainFontDesc = fnt
If Not (MainFont Is Nothing) Then Set MainFont = Nothing
Set MainFont = d3dx.CreateFont(d3ddev, MainFontDesc.hFont)
End Sub

Public Sub DrawText(ByVal Text As String, ByVal Color As Long, ByVal Align As Alignment, ByVal X As Long, ByVal Y As Long)
Dim TextRect As DxVBLiba.RECT
Dim Size As Long
' 1 point = 1.33333333333333 pixels
Size = fnt.Size * 1.33333333333333
TextRect.top = Y
TextRect.Left = X
TextRect.bottom = Y + Int(Size + 3) 'Add three to prevent cutoffs
TextRect.Right = X + Int(Size * Len(Text) + 3) 'Add three to prevent cutoffs
d3dx.DrawText MainFont, Color, Text, TextRect, Align
End Sub

Public Sub DrawSurf(ByRef Source As Direct3DSurface8, ByVal SourceX As Long, ByVal SourceY As Long, ByVal Width As Long, ByVal Height As Long, ByVal Dest As Direct3DSurface8, ByVal DestX As Long, ByVal DestY As Long)
Dim r As DxVBLiba.RECT, p As DxVBLiba.Point

'Set the source rect
r.Left = SourceX * Width
r.top = SourceY * Height
r.Right = r.Left + Width
r.bottom = r.top + Height

'Set the destination point
p.X = DestX
p.Y = DestY

'Draw the surface onto the destination
d3ddev.CopyRects Source, r, 1, Dest, p
End Sub

Public Sub DrawTex(ByRef Source As Direct3DTexture8, PicRect As RECT, PicPosition As D3DVECTOR2, Optional ByVal ScaleSize As Byte = 1, Optional ByVal Alpha As Long = &HFFFFFFFF)
Dim Vect_Scale As D3DVECTOR2
Dim Vect_Rotation As D3DVECTOR2

Vect_Scale.X = ScaleSize
Vect_Scale.Y = ScaleSize
        
Vect_Rotation.X = 0
Vect_Rotation.Y = 0
        
Call d3dsprite.Draw(Source, PicRect, Vect_Scale, Vect_Rotation, 0, PicPosition, Alpha)
End Sub

Public Sub RenderScreen()
d3ddev.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Public Sub BeginScene()
d3ddev.BeginScene
End Sub

Public Sub BeginDrawing()
d3dsprite.Begin
End Sub

Public Sub EndDrawing()
d3dsprite.End
End Sub

Public Sub EndScene()
d3ddev.EndScene
End Sub

Public Sub ClearScreen(ByVal Color As Long)
d3ddev.Clear 0, ByVal 0, D3DCLEAR_TARGET, Color, 1, 0 'Sets screen to color
End Sub

Public Property Get FontSize() As Byte
    FontSize = fnt.Size
End Property

Public Property Get FontName() As Byte
    FontName = fnt.Name
End Property

Public Function ConvertToHex(ByVal lRed As Byte, ByVal lGreen As Byte, ByVal lBlue As Byte, Optional lAlpha As Byte = 0) As Long
Dim ColorHex As String
Dim RedHex As String, BlueHex As String, GreenHex As String, AlphaHex As String

    ' Convert them to hex
    RedHex = CStr(Hex(lRed))
    GreenHex = CStr(Hex(lGreen))
    BlueHex = CStr(Hex(lBlue))
    AlphaHex = CStr(Hex(lAlpha))

    ' Check in case of a single-digit hex (weird, right?)
    If Len(RedHex) = 1 Then RedHex = "0" & RedHex
    If Len(GreenHex) = 1 Then GreenHex = "0" & GreenHex
    If Len(BlueHex) = 1 Then BlueHex = "0" & BlueHex
    If Len(AlphaHex) = 1 Then AlphaHex = "0" & AlphaHex
    
    ' Check in case of an empty string
    If RedHex = "" Then RedHex = "00"
    If GreenHex = "" Then GreenHex = "00"
    If BlueHex = "" Then BlueHex = "00"
    
    ' Set the offical hex
    ColorHex = RedHex & GreenHex & BlueHex

    ' Add the required prefix and the full alpha
    If lAlpha = 0 Then
        ConvertToHex = CLng("&HFF" & ColorHex)
    ' Add the required prefix and the set alpha
    Else
        ColorHex = AlphaHex & ColorHex
        ConvertToHex = CLng("&H" & ColorHex)
    End If
End Function

Private Sub DefineColors()
    ' Define White Color
    C_White = ConvertToHex(255, 255, 255)
    ' Define Black Color
    C_Black = ConvertToHex(0, 0, 0)
    ' Define Blue Color
    C_Blue = ConvertToHex(0, 0, 175)
    ' Define Green Color
    C_Green = ConvertToHex(0, 175, 0)
    ' Define Cyan Color
    C_Cyan = ConvertToHex(0, 175, 175)
    ' Define Red Color
    C_Red = ConvertToHex(175, 0, 0)
    ' Define Magenta Color
    C_Magenta = ConvertToHex(175, 0, 175)
    ' Define Brown Color
    C_Brown = ConvertToHex(100, 50, 0)
    ' Define Grey Color
    C_Grey = ConvertToHex(180, 180, 180)
    ' Define DarkGrey Color
    C_DarkGrey = ConvertToHex(120, 120, 120)
    ' Define BrightBlue Color
    C_BrightBlue = ConvertToHex(0, 0, 255)
    ' Define BrightGreen Color
    C_BrightGreen = ConvertToHex(0, 255, 0)
    ' Define BrightCyan Color
    C_BrightCyan = ConvertToHex(0, 255, 255)
    ' Define BrightRed Color
    C_BrightRed = ConvertToHex(255, 0, 0)
    ' Define Pink Color
    C_Pink = ConvertToHex(255, 0, 175)
    ' Define Yellow Color
    C_Yellow = ConvertToHex(255, 255, 0)
End Sub
