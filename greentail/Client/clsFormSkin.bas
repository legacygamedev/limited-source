Attribute VB_Name = "clsFormSkin"
Option Explicit

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal Y As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long

Private Const RGN_OR = 2
'for movment
'Public Const WM_NCLBUTTONDOWN = &HA1
'Public Const HTCAPTION = 2'

'Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
'Public Declare Sub ReleaseCapture Lib "user32" ()

Function fn_CreateSkin(FormObject As Form, Width As Long, Height As Long, FileName As String, Optional ln_TransColour As Long = 16711935) As Long
On Local Error Resume Next
Dim lRegion As Long
 
   If Dir(FileName) = "" Then
       fn_CreateSkin = 0
       Exit Function
   End If
 
   With FormObject
       .AutoRedraw = True
       .Picture = LoadPicture(FileName, 0)
       .Width = Width * Screen.TwipsPerPixelX
       .Height = Height * Screen.TwipsPerPixelY
       lRegion = fRegionFromBitmap(FormObject, ln_TransColour)
       Call SetWindowRgn(.hWnd, lRegion, True)
   End With
   fn_CreateSkin = 1
 
End Function

Private Function fRegionFromBitmap(picSource As Form, Optional lBackColor As Long) As Long
On Local Error Resume Next
Dim lReturn As Long
Dim lRgnTmp As Long
Dim lSkinRgn As Long
Dim lStart As Long
Dim lRow As Long
Dim lCol As Long
Dim glHeight As Long
Dim glWidth As Long

lSkinRgn = CreateRectRgn(0, 0, 0, 0)
With picSource
   glHeight = .Height / Screen.TwipsPerPixelY
   glWidth = .Width / Screen.TwipsPerPixelX
   If lBackColor < 1 Then lBackColor = GetPixel(.hDC, 0, 0)
   For lRow = 0 To glHeight - 1
       lCol = 0
       Do While lCol < glWidth
           Do While lCol < glWidth And GetPixel(.hDC, lCol, lRow) = lBackColor
               lCol = lCol + 1
           Loop
           If lCol < glWidth Then
               lStart = lCol
               Do While lCol < glWidth And GetPixel(.hDC, lCol, lRow) <> lBackColor
                   lCol = lCol + 1
               Loop
               If lCol > glWidth Then lCol = glWidth
               lRgnTmp = CreateRectRgn(lStart, lRow, lCol, lRow + 1)
               lReturn = CombineRgn(lSkinRgn, lSkinRgn, lRgnTmp, RGN_OR)
               Call DeleteObject(lRgnTmp)
           End If
       Loop
   Next
End With

fRegionFromBitmap = lSkinRgn
End Function

