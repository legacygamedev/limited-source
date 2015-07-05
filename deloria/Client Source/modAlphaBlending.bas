Attribute VB_Name = "modAlphaBlending"
Option Explicit

'For Alpha Blending
Declare Function AlphaBlend Lib "msimg32" (ByVal hDestDC As Long, _
    ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, _
    ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, _
    ByVal HeightSrc As Long, ByVal blendFunct As Long) As Boolean
    
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'alpha belnding
Public Type rBlendProps
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type

Public Sub SquareAlphaBlend( _
ByVal cSrc_Widht As Integer, _
ByVal cSrc_Height As Integer, _
ByVal cSrc As Long, _
ByVal cSrc_X As Integer, _
ByVal cSrc_Y As Integer, _
ByVal cDest As Long, _
ByVal cDest_X As Integer, _
ByVal cDest_Y As Integer, _
ByVal nLevel As Byte)

    Dim LrProps As rBlendProps
    Dim LnBlendPtr As Long
    
    LrProps.tBlendAmount = nLevel
    CopyMemory LnBlendPtr, LrProps, 4
    
    AlphaBlend cDest, cDest_X, cDest_Y, cSrc_Widht, cSrc_Height, cSrc, cSrc_X, cSrc_Y, cSrc_Widht, cSrc_Height, LnBlendPtr
    
End Sub

