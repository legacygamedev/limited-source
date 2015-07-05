Attribute VB_Name = "modGameIcon"
Option Explicit

' /* CREDITS: http://www.vbaccelerator.com/home/VB/Tips/Setting_the_App_Icon_Correctly/article.asp */
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Const SM_CXICON = 11

Private Const SM_CYICON = 12

Private Const SM_CXSMICON = 49

Private Const SM_CYSMICON = 50

Private Declare Function LoadImageAsString _
                Lib "user32" _
                Alias "LoadImageA" (ByVal hInst As Long, _
                                    ByVal lpsz As String, _
                                    ByVal uType As Long, _
                                    ByVal cxDesired As Long, _
                                    ByVal cyDesired As Long, _
                                    ByVal fuLoad As Long) As Long

Private Const LR_DEFAULTCOLOR = &H0

Private Const LR_MONOCHROME = &H1

Private Const LR_COLOR = &H2

Private Const LR_COPYRETURNORG = &H4

Private Const LR_COPYDELETEORG = &H8

Private Const LR_LOADFROMFILE = &H10

Private Const LR_LOADTRANSPARENT = &H20

Private Const LR_DEFAULTSIZE = &H40

Private Const LR_VGACOLOR = &H80

Private Const LR_LOADMAP3DCOLORS = &H1000

Private Const LR_CREATEDIBSECTION = &H2000

Private Const LR_COPYFROMRESOURCE = &H4000

Private Const LR_SHARED = &H8000&

Private Const IMAGE_ICON = 1

Private Declare Function SendMessageLong _
                Lib "user32" _
                Alias "SendMessageA" (ByVal hWnd As Long, _
                                      ByVal wMsg As Long, _
                                      ByVal wParam As Long, _
                                      ByVal lParam As Long) As Long

Private Const WM_SETICON = &H80

Private Const ICON_SMALL = 0

Private Const ICON_BIG = 1

Private Declare Function GetWindow _
                Lib "user32" (ByVal hWnd As Long, _
                              ByVal wCmd As Long) As Long

Private Const GW_OWNER = 4

Public Sub SetIcon()

    Dim lhWndTop   As Long

    Dim lhWnd      As Long

    Dim cx         As Long

    Dim cy         As Long

    Dim hIconLarge As Long

    Dim hIconSmall As Long
      
    ' Find VB's hidden parent window:

    lhWnd = frmMain.hWnd
    lhWndTop = lhWnd

    Do While Not (lhWnd = 0)
        lhWnd = GetWindow(lhWnd, GW_OWNER)

        If Not (lhWnd = 0) Then
            lhWndTop = lhWnd
        End If

    Loop
   
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
    hIconLarge = LoadImageAsString(vbNull, App.Path & "\64x64.ico", IMAGE_ICON, cx, cy, LR_LOADFROMFILE)
    SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    SendMessageLong frmMain.hWnd, WM_SETICON, ICON_BIG, hIconLarge
   
    cx = GetSystemMetrics(SM_CXSMICON)
    cy = GetSystemMetrics(SM_CYSMICON)
    hIconSmall = LoadImageAsString(0, App.Path & "\16x16.ico", IMAGE_ICON, 16, 16, LR_SHARED)
    SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    SendMessageLong frmMain.hWnd, WM_SETICON, ICON_SMALL, hIconSmall
      
End Sub




