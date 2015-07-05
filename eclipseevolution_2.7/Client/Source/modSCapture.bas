Attribute VB_Name = "modSCapture"
Option Explicit

Public Type PicBmp
    Size As Long
    Type As Long
        hBmp As Long
        hPal As Long
        Reserved As Long
    End Type

    Public Type PALETTEENTRY
        peRed As Byte
        peGreen As Byte
        peBlue As Byte
        peFlags As Byte
    End Type

    Public Type LOGPALETTE
        palVersion As Integer
        palNumEntries As Integer
        palPalEntry(255) As PALETTEENTRY ' Enough for 256 colors.
    End Type

    Public Type GUID
        Data1 As Long
        Data2 As Integer
        Data3 As Integer
        Data4(7) As Byte
    End Type

    Public Const RASTERCAPS As Long = 38
    Public Const RC_PALETTE As Long = &H100
    Public Const SIZEPALETTE As Long = 104

    Public Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
    Public Declare Function CreateCompatibleBitmap Lib "gdi32.dll" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Attribute CreateCompatibleBitmap.VB_UserMemId = 1879048232
    Public Declare Function GetDeviceCaps Lib "gdi32.dll" (ByVal hDC As Long, ByVal iCapabilitiy As Long) As Long
Attribute GetDeviceCaps.VB_UserMemId = 1879048276
    Public Declare Function GetSystemPaletteEntries Lib "gdi32.dll" (ByVal hDC As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Attribute GetSystemPaletteEntries.VB_UserMemId = 1879048312
    Public Declare Function CreatePalette Lib "gdi32.dll" (lpLogPalette As LOGPALETTE) As Long
Attribute CreatePalette.VB_UserMemId = 1879048356
    Public Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Attribute SelectObject.VB_UserMemId = 1879048392
    Public Declare Function DeleteDC Lib "gdi32.dll" (ByVal hDC As Long) As Long
Attribute DeleteDC.VB_UserMemId = 1879048428
    Public Declare Function GetForegroundWindow Lib "USER32.DLL" () As Long
Attribute GetForegroundWindow.VB_UserMemId = 1879048460
    Public Declare Function SelectPalette Lib "gdi32.dll" (ByVal hDC As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Attribute SelectPalette.VB_UserMemId = 1879048500
    Public Declare Function RealizePalette Lib "gdi32.dll" (ByVal hDC As Long) As Long
Attribute RealizePalette.VB_UserMemId = 1879048536
    Public Declare Function GetWindowDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Attribute GetWindowDC.VB_UserMemId = 1879048572
    Public Declare Function GetDC Lib "USER32.DLL" (ByVal hWnd As Long) As Long
Attribute GetDC.VB_UserMemId = 1879048604
    Public Declare Function GetWindowRect Lib "USER32.DLL" (ByVal hWnd As Long, lpRect As RECT) As Long
Attribute GetWindowRect.VB_UserMemId = 1879048632
    Public Declare Function ReleaseDC Lib "USER32.DLL" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Attribute ReleaseDC.VB_UserMemId = 1879048668
    Public Declare Function GetDesktopWindow Lib "USER32.DLL" () As Long
Attribute GetDesktopWindow.VB_UserMemId = 1879048700
    Public Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
Attribute OleCreatePictureIndirect.VB_UserMemId = 1879048740

Public Function CaptureScreen() As Picture
    Dim hWndScreen As Long
    On Error Resume Next

    hWndScreen = GetDesktopWindow()
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function

Public Function CaptureForm(frmSrc As Form) As Picture
    On Error Resume Next

    Set CaptureForm = CaptureWindow(frmSrc.hWnd, False, 0, 0, frmSrc.ScaleX(frmSrc.Width, vbTwips, vbPixels), frmSrc.ScaleY(frmSrc.Height, vbTwips, vbPixels))
End Function

Public Function CaptureClient(frmSrc As Form) As Picture
    On Error Resume Next

    Set CaptureClient = CaptureWindow(frmSrc.hWnd, True, 0, 0, frmSrc.ScaleX(frmSrc.ScaleWidth, frmSrc.ScaleMode, vbPixels), frmSrc.ScaleY(frmSrc.ScaleHeight, frmSrc.ScaleMode, vbPixels))
End Function

Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    On Error Resume Next

    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID

    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With

    With Pic
        .Size = Len(Pic) ' Length of structure.
        .Type = vbPicTypeBitmap ' Type of Picture (bitmap).
        .hBmp = hBmp ' Handle to bitmap(GetPlayerMap(myindex)).
        .hPal = hPal ' Handle to palette (may be null).
    End With

    OleCreatePictureIndirect Pic, IID_IDispatch, 1, IPic

    Set CreateBitmapPicture = IPic
End Function

Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    On Error Resume Next

    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE

    If Client Then
        hDCSrc = GetDC(hWndSrc) ' Get device context for client area.
    Else
        hDCSrc = GetWindowDC(hWndSrc) ' Get device context for entire window.
    End If

    hDCMemory = CreateCompatibleDC(hDCSrc)

    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS) ' Raster capabilities.
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE ' Palette support.
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE) ' Size of palette.

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then

        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        GetSystemPaletteEntries hDCSrc, 0, 256, LogPal.palPalEntry(0)
        hPal = CreatePalette(LogPal)

        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        RealizePalette hDCMemory
    End If

    BitBlt hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy

    hBmp = SelectObject(hDCMemory, hBmpPrev)

    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If

    DeleteDC hDCMemory
    ReleaseDC hWndSrc, hDCSrc

    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function

Public Function CaptureArea(frmSrc As Form, Left As Long, Top As Long, Width As Long, Height As Long) As Picture
    On Error Resume Next

    Set CaptureArea = CaptureWindow(frmSrc.hWnd, True, Left, Top, Width, Height)
End Function

Public Function CaptureActiveWindow() As Picture
    Dim hWndActive As Long
    Dim RectActive As RECT
    On Error Resume Next

    hWndActive = GetForegroundWindow()

    GetWindowRect hWndActive, RectActive

    Set CaptureActiveWindow = CaptureWindow(hWndActive, False, 0, 0, RectActive.Right - RectActive.Left, RectActive.Bottom - RectActive.Top)
End Function

