Attribute VB_Name = "SurfUtil"
'************************************************
'*  SurfUtil.bas                                *
'*                                              *
'* By: W-Buffer                                 *
'* Web: istudios.virtualave.net/vb/             *
'* Mail: wbuffer@hotmail.com                    *
'*                                              *
'* Modified by: Don Andy (don_andy@gmx.de)      *
'*                                              *
'* Notes: Do whatever you want with this bas    *
'*        (Steal, Copy, Etc.)                   *
'*        These functions were modified to work *
'*        with PNG only and to recieve the DDSD *
'*        The DDSD MUST have a DDSD_HEIGHT and  *
'*        DDSD_WIDTH flag!!                     *
'*        The lib used to display PNGs is the   *
'*        PaintX-Lib (http://www.paintlib.de)   *
'************************************************

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Function LoadPNG(FileName As String) As StdPicture
    Dim PictureDecoder As New PAINTXLib.PictureDecoder
    Set LoadPNG = PictureDecoder.LoadPicture(FileName)
End Function

Public Function LoadImage(FileName As String, DDraw As DirectDraw7, SDesc As DDSURFACEDESC2) As DirectDrawSurface7
    Dim TPict As StdPicture
    Set TPict = LoadPNG(FileName)
    
    SDesc.lHeight = CLng((TPict.Height * 0.001) * 567 / Screen.TwipsPerPixelY)
    SDesc.lWidth = CLng((TPict.Width * 0.001) * 567 / Screen.TwipsPerPixelX)
    
    
    Set LoadImage = DDraw.CreateSurface(SDesc)
    
    Dim SDC As Long, TDC As Long
    SDC = LoadImage.GetDC
    TDC = CreateCompatibleDC(0)
    SelectObject TDC, TPict.Handle
    
    BitBlt SDC, 0, 0, SDesc.lWidth, SDesc.lHeight, TDC, 0, 0, vbSrcCopy
        
    LoadImage.ReleaseDC SDC
    DeleteDC TDC
       
    Set TPict = Nothing
End Function

Public Function LoadImageStretch(FileName As String, Height As Long, Width As Long, DDraw As DirectDraw7, SDesc As DDSURFACEDESC2) As DirectDrawSurface7
    Dim TPict As New StdPicture
    Set TPict = LoadPNG(FileName)
    
    SDesc.lHeight = Height
    SDesc.lWidth = Width
    
    Set LoadImageStretch = DDraw.CreateSurface(SDesc)
    
    Dim SDC As Long, TDC As Long
    SDC = LoadImageStretch.GetDC
    TDC = CreateCompatibleDC(0)
    SelectObject TDC, TPict.Handle
    
    StretchBlt SDC, 0, 0, Width, Height, TDC, 0, 0, CLng((TPict.Width * 0.001) * 567 / Screen.TwipsPerPixelX), CLng((TPict.Height * 0.001) * 567 / Screen.TwipsPerPixelY), vbSrcCopy
    
    LoadImageStretch.ReleaseDC SDC
    DeleteDC TDC
        
    Set TPict = Nothing
End Function
