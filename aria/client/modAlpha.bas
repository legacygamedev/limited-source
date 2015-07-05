Attribute VB_Name = "modAlpha"
Option Explicit
' //========= ALPHA LIBRARY ==========\\
'    CREDS TO EMB FOR THIS, IM UBER!!!
' \\==================================//

'Function call for accessing the vbDABL library:
Public Declare Function vbDABLalphablend16 Lib "vbDABL" (ByVal iMode As Integer, ByVal bColorKey As Integer, ByRef sPtr As Any, ByRef dPtr As Any, ByVal iAlphaVal As Integer, ByVal iWidth As Integer, ByVal iHeight As Integer, ByVal isPitch As Integer, ByVal idPitch As Integer, ByVal iColorKey As Integer) As Integer

'Creds to the guy who made the vbDABL Library and the guy who made this code who I lost the link to... (Modified by me :))
Public Sub AlphaBlend(DD_Src As DirectDrawSurface7, srcRect As RECT, DD_Dst As DirectDrawSurface7, x As Long, y As Long, alphaval As Long)
'On Error Resume Next

    Dim tempDDSD As DDSURFACEDESC2
    Dim temp2DDSD As DDSURFACEDESC2
    Dim RECTvar As RECT
    Dim ddsBackArray() As Byte
    Dim ddsForeArray() As Byte
    Dim emptyrect As RECT
    
    RECTvar = srcRect
    
    If RECTvar.Right < RECTvar.Left + 3 Then
        'don't draw anything, quit
        Exit Sub
    End If
    
    If RECTvar.Bottom < RECTvar.Top + 3 Then
        Exit Sub
    End If
    
    'Pass a empty rec, locks the whole surface.
    DD_Dst.Lock emptyrect, temp2DDSD, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
    DD_Src.Lock srcRect, tempDDSD, DDLOCK_NOSYSLOCK Or DDLOCK_WAIT, 0
    
    DD_Dst.GetLockedArray ddsBackArray
    DD_Src.GetLockedArray ddsForeArray
    
    Select Case DDSD_BackBuffer.ddpfPixelFormat.lGBitMask
        Case &H3E0 '555 mode
            vbDABLalphablend16 555, 1, ddsForeArray(RECTvar.Left + RECTvar.Left, RECTvar.Top), ddsBackArray(x + x, y), alphaval, (RECTvar.Right - RECTvar.Left), (RECTvar.Bottom - RECTvar.Top), tempDDSD.lPitch, temp2DDSD.lPitch, 0
        Case &H7E0 '565 mode
            vbDABLalphablend16 565, 1, ddsForeArray(RECTvar.Left + RECTvar.Left, RECTvar.Top), ddsBackArray(x + x, y), alphaval, (RECTvar.Right - RECTvar.Left), (RECTvar.Bottom - RECTvar.Top), tempDDSD.lPitch, temp2DDSD.lPitch, 0
        Case Else 'Wth? Assume 555
            vbDABLalphablend16 555, 1, ddsForeArray(RECTvar.Left + RECTvar.Left, RECTvar.Top), ddsBackArray(x + x, y), alphaval, (RECTvar.Right - RECTvar.Left), (RECTvar.Bottom - RECTvar.Top), tempDDSD.lPitch, temp2DDSD.lPitch, 0
    End Select
    
    DD_Src.Unlock srcRect
    DD_Dst.Unlock emptyrect
End Sub

