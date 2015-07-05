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

'//The variables Required
Dim Dx As DirectX8 'The master Object, everything comes from here
Dim D3D As Direct3D8 'This controls all things 3D
Dim D3DDevice As Direct3DDevice8 'This actually represents the hardware doing the rendering
Dim D3DX As D3DX8 '//A helper library
Dim Texture As Direct3DTexture8
Dim Surface As Direct3DSurface8
'Dim bRunning As Boolean 'Controls whether the program is running or not...

'This is the Flexible-Vertex-Format description for a 2D vertex (Transformed and Lit)
Const FVF = D3DFVF_XYZRHW Or D3DFVF_TEX1 Or D3DFVF_DIFFUSE Or D3DFVF_SPECULAR

'This structure describes a transformed and lit vertex - it's identical to the DirectX7 type "D3DTLVERTEX"
Private Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    color As Long
    specular As Long
    tu As Single
    tv As Single
End Type

Dim TriStrip(0 To 3) As TLVERTEX '//We need 4 vertices to make a simple square...
Dim TriList(0 To 5) As TLVERTEX '//Note that, like in TriStrip, we're generating 2 triangles - yet using 2 more vertices...
Dim TriFan(0 To 5) As TLVERTEX

Private Sub Class_Initialize()
    If Initialize = False Then
        MsgBox "Error!"
    End If
End Sub

Public Function Initialize()
Dim ColorKeyVal As Long
'##################################################
'## FIRST WE NEED TO ACTUALLY CREATE THE OBJECT
'##################################################

Dim DispMode As D3DDISPLAYMODE '//Describes our Display Mode
Dim D3DWindow As D3DPRESENT_PARAMETERS '//Describes our Viewport

Set Dx = New DirectX8 '//Create our Master Object
Set D3D = Dx.Direct3DCreate() '//Make our Master Object create the Direct3D Interface

Set D3DX = New D3DX8 '//Create our helper library...

'DispMode.Format = D3DFMT_X8R8G8B8
DispMode.Format = D3DFMT_R5G6B5 'If this mode doesn't work try the commented one above...

D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode '//Retrieve the current display Mode

D3DWindow.Windowed = 1 '//Tell it we're using Windowed Mode
'D3DWindow.SwapEffect = D3DSWAPEFFECT_COPY_VSYNC '//We'll refresh when the monitor does
D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP
D3DWindow.BackBufferFormat = DispMode.Format '//We'll use the format we just retrieved...
D3DWindow.hDeviceWindow = frmMain.hWnd

'//This line will be explained in detail in a minute...
Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.hWnd, _
                                                     D3DCREATE_SOFTWARE_VERTEXPROCESSING, _
                                                     D3DWindow)

'########################################################
'## THEN WE'LL NEED TO CONFIGURE THE RENDERING DEVICE
'########################################################

'//Set the vertex shader to use our vertex format
D3DDevice.SetVertexShader FVF

'//Transformed and lit vertices dont need lighting
'   so we disable it...
D3DDevice.SetRenderState D3DRS_LIGHTING, False

'//Set alpha blending to true
D3DDevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
D3DDevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
D3DDevice.SetRenderState D3DRS_ALPHABLENDENABLE, True

'//Choose one of the following depending on what you
'   should need. Other colours can be made up, but these
'   ones should be okay for most uses...

'ColorKeyVal = &HFF000000 '//Black
'ColorKeyVal = &HFFFF0000 '//Red
'ColorKeyVal = &HFF00FF00 '//Green
'ColorKeyVal = &HFF0000FF '//Blue
'ColorKeyVal = &HFFFF00FF '//Magenta
'ColorKeyVal = &HFFFFFF00 '//Yellow
'ColorKeyVal = &HFF00FFFF '//Cyan
'ColorKeyVal = &HFFFFFFFF '//White
ColorKeyVal = &HF00B7EF

'//We now want to load our texture;
'Set Texture = D3DX.CreateTextureFromFile(D3DDevice, App.Path & "\Graphics\Sprites.bmp")

'Open App.Path & "\Graphics\Sprites.bmp" For Binary As #1
'    Dim BinaryFile(LOF(1) - 1)
'    Get #1, , BinaryFile()
'    Set Texture = D3DX.CreateTextureFromFileInMemoryEx(D3DDevice, BinaryFile(0), LOF(1), D3DX_DEFAULT, D3DX_DEFAULT, _
'                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
'                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
'                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
'                                                                            ByVal 0, ByVal 0)
'Close #1

Set Texture = D3DX.CreateTextureFromFileEx(D3DDevice, App.Path & "\Graphics\Sprite.bmp", D3DX_DEFAULT, D3DX_DEFAULT, _
                                                                            D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, _
                                                                            D3DPOOL_MANAGED, D3DX_FILTER_POINT, _
                                                                            D3DX_FILTER_POINT, ColorKeyVal, _
                                                                            ByVal 0, ByVal 0)



'############################################
'## FINALLY WE'LL IMPLEMENT A NEW FUNCTION
'############################################

'//We can only continue if Initialize Geometry succeeds;
'   If it doesn't we'll fail this call as well...
If InitializeGeometry() = True Then
    Initialize = True '//We succeeded
    Exit Function
End If

End Function

'//This basically just fills out the relevent array of vertices
'with the information relevent to what we want to display - A square!
Private Function InitializeGeometry() As Boolean
    '//Geometry is basically vertices; you'll find it being
    '   used to represent vertices/triangles/meshs as we go along
    
On Error GoTo BailOut: '//Setup our Error handler
'##################
'## TRIANGLE STRIP
'#################

            'We're going to create a square (easily adapted to be a rectangle)
            'by defining 4 vertices in a particular order - ClockWise.
            'If you define them in the wrong order they either wont appear at all
            'or wont appear as you intended....
            
            '0 - - - - - 1
            '             /
            '           /
            '         /
            '       /
            '     /
            '   /
            ' /
            '2 - - - - - 3
            
            '//Ignore the ' parts down the side... Only the Z shape should be there...
            
            '//As you can see this is basically two triangles.
            
            'vertex 0
            'TriStrip(0) = CreateTLVertex(10, 10, 0, 1, RGB(255, 255, 255), 0, 0, 0)
            
            'vertex 1
            'TriStrip(1) = CreateTLVertex(210, 10, 0, 1, RGB(255, 255, 255), 0, 1, 0)
            
            'vertex 2
            'TriStrip(2) = CreateTLVertex(10, 210, 0, 1, RGB(255, 255, 255), 0, 0, 1)
            
            'vertex 3
            'TriStrip(3) = CreateTLVertex(210, 210, 0, 1, RGB(255, 255, 255), 0, 1, 1)
            
            'vertex 0
            TriStrip(0) = CreateTLVertex(10, 10, 0, 1, RGB(255, 255, 255), 0, 0, 0)
            
            'vertex 1
            TriStrip(1) = CreateTLVertex(42, 10, 0, 1, RGB(255, 255, 255), 0, 1, 0)
            
            'vertex 2
            TriStrip(2) = CreateTLVertex(10, 42, 0, 1, RGB(255, 255, 255), 0, 0, 1)
            
            'vertex 3
            TriStrip(3) = CreateTLVertex(42, 42, 0, 1, RGB(255, 255, 255), 0, 1, 1)


'#################
'## TRIANGLE LIST
'################

'We're going to make an arrow type shape from two triangles; whilst this is still perfectly possible
'using a triangle strip - it's a good enough example of using triangle lists....


    '// TRIANGLE 1
    
    'vertex 0
    TriList(0) = CreateTLVertex(210, 210, 0, 1, RGB(0, 0, 0), 0, 0, 0)
    
    'vertex 1
    TriList(1) = CreateTLVertex(400, 250, 0, 1, RGB(255, 255, 255), 0, 0, 0)
    
    'vertex 2
    TriList(2) = CreateTLVertex(250, 250, 0, 1, RGB(255, 255, 0), 0, 0, 0)
    
    '//TRIANGLE 2
    
    'vertex 3
    TriList(3) = CreateTLVertex(210, 210, 0, 1, RGB(0, 255, 255), 0, 0, 0)
    
    'vertex 4
    TriList(4) = CreateTLVertex(250, 250, 0, 1, RGB(255, 0, 255), 0, 0, 0)
    
    'vertex 5
    TriList(5) = CreateTLVertex(250, 400, 0, 1, RGB(0, 255, 0), 0, 0, 0)
    
    
'#################
'## TRIANGLE FAN
'################

'We're now going to set up a triangle fan; these are best used for circular
'type shapes. We'll just generate a hexagonal type shape..

    '//MAIN POINT
    TriFan(0) = CreateTLVertex(300, 100, 0, 1, RGB(255, 0, 0), 0, 0, 0)
    
    '//Subsequent vertices around the shape
    TriFan(1) = CreateTLVertex(350, 50, 0, 1, RGB(0, 255, 0), 0, 0, 0)
    TriFan(2) = CreateTLVertex(450, 50, 0, 1, RGB(0, 0, 255), 0, 0, 0)
    TriFan(3) = CreateTLVertex(500, 100, 0, 1, RGB(255, 0, 255), 0, 0, 0)
    TriFan(4) = CreateTLVertex(450, 150, 0, 1, RGB(255, 255, 0), 0, 0, 0)
    TriFan(5) = CreateTLVertex(350, 150, 0, 1, RGB(255, 255, 255), 0, 0, 0)

InitializeGeometry = True
Exit Function
BailOut:
InitializeGeometry = False
End Function

'//This is just a simple wrapper function that makes filling the structures much much easier...
Private Function CreateTLVertex(X As Single, Y As Single, Z As Single, rhw As Single, color As Long, _
                                               specular As Long, tu As Single, tv As Single) As TLVERTEX
    '//NB: whilst you can pass floating point values for the coordinates (single)
    '       there is little point - Direct3D will just approximate the coordinate by rounding
    '       which may well produce unwanted results....
CreateTLVertex.X = X
CreateTLVertex.Y = Y
CreateTLVertex.Z = Z
CreateTLVertex.rhw = rhw
CreateTLVertex.color = color
CreateTLVertex.specular = specular
CreateTLVertex.tu = tu
CreateTLVertex.tv = tv
End Function

Public Sub SetRenderTex()
    D3DDevice.SetTexture 0, Texture '//Tell the device which texture we want to use...
    D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
End Sub

Public Sub Render()
'//1. We need to clear the render device before we can draw anything
'       This must always happen before you start rendering stuff...
D3DDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, RGB(0, 255, 0), 1#, 0 '//Clear the screen black

'//2. Rendering the graphics...

D3DDevice.BeginScene
    'All rendering calls go between these two lines
    
    'Comment out some of the lines below should the screen seem
    'to clustered
    
    'First the Triangle Strip
    'D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))
    'D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(1), Len(TriStrip(1)) '//Nothing happens
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
    'D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG2, D3DTA_CURRENT
    D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TriStrip(0), Len(TriStrip(0))
    
    'Next the triangle list
    'D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLELIST, 2, TriList(0), Len(TriList(0))
    
    'Finally, the triangle fan
    'D3DDevice.DrawPrimitiveUP D3DPT_TRIANGLEFAN, 4, TriFan(0), Len(TriFan(0))
    
D3DDevice.EndScene

'//3. Update the frame to the screen...
'       This is the same as the Primary.Flip method as used in DirectX 7
'       These values below should work for almost all cases...
D3DDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
End Sub

Private Sub Class_Terminate()
    'MsgBox "Unload Direct3D"
    Set D3DDevice = Nothing
    Set D3D = Nothing
    Set Dx = Nothing
End Sub

Private Function TexCoord(ByVal MaxTexSize As Long, ByVal PixelCoord As Long) As Single
    TexCoord = (1 / MaxTexSize) * PixelCoord
End Function

Public Sub SetTriStripTopLeft(ByVal X As Single, ByVal Y As Single, ByVal Start_X As Single, ByVal Start_Y As Single, ByVal XL As Single, ByVal YL As Single, Optional Alpha As String = "&HFFFFFFFF")
Dim lAlpha As Long
lAlpha = CLng(Alpha)
lAlpha = CLng("&HFFFFFFFF")
    'vertex 0
    TriStrip(0) = CreateTLVertex(X, Y, 0, 1, lAlpha, 0, 0, 0)
lAlpha = CLng("&HFFFFFFFF")
    'vertex 1
    TriStrip(1) = CreateTLVertex(X + XL, Y, 0, 1, lAlpha, 0, 1, 0)
lAlpha = CLng("&HFFFFFFFF")
    'vertex 2
    TriStrip(2) = CreateTLVertex(X, Y + YL, 0, 1, lAlpha, 0, 0, 1)
lAlpha = CLng("&HFFFFFFFF")
    'vertex 3
    TriStrip(3) = CreateTLVertex(X + XL, Y + YL, 0, 1, lAlpha, 0, 1, 1)
End Sub
