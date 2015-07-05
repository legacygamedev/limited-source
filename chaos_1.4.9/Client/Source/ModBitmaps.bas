Attribute VB_Name = "ModBitmaps"
Option Explicit

Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function BitBlt& Lib "gdi32" (ByVal hDestDC&, ByVal X&, ByVal Y&, ByVal nWidth&, ByVal nHeight&, ByVal hSrcDC&, ByVal xSrc&, ByVal ySrc&, ByVal dwRop&)
Private Declare Function SetBkColor& Lib "gdi32" (ByVal hDC&, ByVal crColor&)
Public Declare Function GetPixel& Lib "gdi32" (ByVal hDC&, ByVal X&, ByVal Y&)
Private Declare Function CreateCompatibleBitmap& Lib "gdi32" (ByVal hDC&, ByVal nWidth&, ByVal nHeight&)
Private Declare Function CreateCompatibleDC& Lib "gdi32" (ByVal hDC&)
Private Declare Function SelectObject& Lib "gdi32" (ByVal hDC&, ByVal hObject&)
Private Declare Function DeleteObject& Lib "gdi32" (ByVal hObject&)
Private Declare Function CreateBitmap& Lib "gdi32" (ByVal nWidth&, ByVal nHeight&, ByVal nPlanes&, ByVal nBitCount&, ByVal lpBits As Any)
Private Declare Function DeleteDC& Lib "gdi32" (ByVal hDC&)

Private Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private Const SRCINVERT = &H660046   ' (DWORD) dest = source XOR dest
Public Sub BitBltItNow(hDestDC As Long, lDestX, lDestY, hSourceDC As Long, lSourceW As Long, lSourceH As Long, lStartX As Long, lStartY As Long)
    'Just a plain old BitBlt
    
    'hDestDC = were the image will be copied to
    'lDestX & lDestY = the X and Y coordinates were it will end up
    'lSourceW &  lSourceH = the width and height of the bitmap copied
    'lSourceH = the source bitmap
    'lStartX & lStartY = the X and Y coordinates of were to start copying on the source
    
    Dim lRet As Long
    lRet = BitBlt(hDestDC, lDestX, lDestY, lSourceW, lSourceH, hSourceDC, lStartX, lStartY, SRCCOPY)
End Sub

Public Sub StretchBitMap(hDestDC As Long, lDestX As Long, lDestY As Long, lDestW As Long, lDestH As Long, hSourceDC As Long, lSourceW As Long, lSourceH As Long)
    'This stretches source bitmap
    
    'The only difference between this and BitBltItNow is the lSourceW, lSourceH arguments
    'If they are larger than lDestW, lDestH the bitmap will be stretched
    'If they are smaller than lDestW, lDestH the bitmap will be shrunk
    Call StretchBlt(hDestDC, lDestX, lDestY, lDestW, lDestH, hSourceDC, 0, 0, lSourceW, lSourceH, SRCCOPY)
End Sub

Public Sub TransTileToForm(hDestDC As Long, lDestW As Long, lDestH As Long, hSourceDC As Long, lSourceW As Long, lSourceH As Long, lTransColor As Long)
        
    'Same thing as TileToForm except that it calls TransBltNow instead of a basic BitBlt
    Dim lRet As Long
    Dim lRows As Long
    Dim lCols As Long
    Dim I As Long
    Dim J As Long
    Dim lDestX As Long
    Dim lDestY As Long
    
    lCols = lDestW \ lSourceW
    lRows = lDestH \ lSourceH

    For I = 0 To lCols
        lDestX = I * lSourceW
        For J = 0 To lRows
            lDestY = J * lSourceH
            TransBltNow hDestDC, lDestX, lDestY, lSourceW, lSourceH, hSourceDC, 0, 0, lTransColor
        Next
    Next
End Sub
Public Sub TileToForm(hDestDC As Long, lDestW As Long, lDestH As Long, hSourceDC As Long, lSourceW As Long, lSourceH As Long)

    'Tiles to source bitmap on to the destination
    Dim lRet As Long
    Dim lRows As Long
    Dim lCols As Long
    Dim I As Long
    Dim J As Long
    Dim lDestX As Long
    Dim lDestY As Long
    
    'Figure out how many bitmaps will fit across
    lCols = lDestW \ lSourceW
    'Figure out how many bitmaps will fit down
    lRows = lDestH \ lSourceH

    'A nested loop to copy rows and cols
    For I = 0 To lCols
        lDestX = I * lSourceW
        For J = 0 To lRows
            lDestY = J * lSourceH
            lRet = BitBlt(hDestDC, lDestX, lDestY, lSourceW, lSourceH, hSourceDC, 0, 0, SRCCOPY)
        Next
    Next
End Sub


Public Sub TransBltNow(hDestDC As Long, lDestX As Long, lDestY As Long, lWidth As Long, lHeight As Long, hSourceDC As Long, lSourceX As Long, lSourceY As Long, lTransColor As Long)
'   This function copies a bitmap from one device context to the other
'   where every pixel in the source bitmap that matches the specified color
'   becomes transparent, letting the destination bitmap show through.

    Dim lOldColor As Long
    Dim hMaskDC As Long
    Dim hMaskBmp As Long
    Dim hOldMaskBmp As Long
    Dim hTempBmp As Long
    Dim hTempDC As Long
    Dim hOldTempBmp As Long
    Dim hDummy As Long
    Dim lRet As Long

    '   The Background colors of Source and Destination DCs must
    '   be the transparancy color in order to create a mask.
    lOldColor = SetBkColor&(hSourceDC, lTransColor)
    lOldColor = SetBkColor&(hDestDC, lTransColor)
    
    '   The mask DC must be compatible with the destination dc,
    '   but the mask has to be created as a monochrome bitmap.
    '   For this reason, we create a compatible dc and bitblt
    '   the mono mask into it.
    
    '   Create the Mask DC, and a compatible bitmap to go in it.
    hMaskDC = CreateCompatibleDC(hDestDC)
    hMaskBmp = CreateCompatibleBitmap(hDestDC, lWidth, lHeight)
    '   Move the Mask bitmap into the Mask DC
    hOldMaskBmp = SelectObject(hMaskDC, hMaskBmp)
    
    '   Create a monochrome bitmap that will be the actual mask bitmap.
    hTempBmp = CreateBitmap(lWidth, lHeight, 1, 1, 0&)
    '   Create a temporary DC, and put the mono bitmap into it
    hTempDC = CreateCompatibleDC(hDestDC)
    hOldTempBmp = SelectObject(hTempDC, hTempBmp)
    
    '   BitBlt the Source image into the mono dc to create a mono mask.
    If BitBlt(hTempDC, 0, 0, lWidth, lHeight, hSourceDC, lSourceX, lSourceY, SRCCOPY) Then
        '   Copy the mono mask into our Mask DC
        hDummy = BitBlt(hMaskDC, 0, 0, lWidth, lHeight, hTempDC, 0, 0, SRCCOPY)
    End If
    
    '   Clean up temp DC and bitmap
    hTempBmp = SelectObject(hTempDC, hOldTempBmp)
    hDummy = DeleteObject(hTempBmp)
    hDummy = DeleteDC(hTempDC)
    
    '   Copy the source to the destination with XOR
    lRet = BitBlt(hDestDC, lDestX, lDestY, lWidth, lHeight, hSourceDC, lSourceX, lSourceY, SRCINVERT)
    '   Copy the Mask to the destination with AND
    lRet = BitBlt(hDestDC, lDestX, lDestY, lWidth, lHeight, hMaskDC, 0, 0, SRCAND)
    '   Again, copy the source to the destination with XOR
    lRet = BitBlt(hDestDC, lDestX, lDestY, lWidth, lHeight, hSourceDC, lSourceX, lSourceY, SRCINVERT)

    '   Clean up mask DC and bitmap
    hMaskBmp = SelectObject(hMaskDC, hOldMaskBmp)
    hDummy = DeleteObject(hMaskBmp)
    hDummy = DeleteDC(hMaskDC)

End Sub


