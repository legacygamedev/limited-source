VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cAniCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' FileName:    cAniCursor
' Author:      Steve McMahon
' Date:        20 March 1999
'
' Loads and maintains an animated or standard cursor
' from a resource or file and allows the frames to
' be drawn onto a DC.
'
' ---------------------------------------------------------------------------
' Visit vbAccelerator at
'     http://vbaccelerator.com
' ===========================================================================

Private Declare Function DrawIconEx Lib "user32" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function LoadImageLong Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long
Private Declare Function LoadImageString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cx As Long, ByVal cy As Long, ByVal uFlags As Long) As Long

Private Const IMAGE_BITMAP = 0&
Private Const IMAGE_ICON = 1&
Private Const IMAGE_CURSOR = 2&

Private Const LR_LOADMAP3DCOLORS = &H1000&
Private Const LR_LOADFROMFILE = &H10&
Private Const LR_LOADTRANSPARENT = &H20&

Private Const DI_MASK = &H1&
Private Const DI_IMAGE = &H2&
Private Const DI_NORMAL = &H3&
Private Const DI_COMPAT = &H4&
Private Const DI_DEFAULTSIZE = &H8&

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long


' To Report API errors:
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GCL_HCURSOR = (-12)

Private m_hCur As Long
Private m_iFrame As Long
Private m_iNumFrames As Long

Private Type tClassCursor
   hWnd As Long
   hCurOld As Long
End Type
Private m_tClassCursor() As tClassCursor
Private m_iCount As Long
Private m_bInDevelopment As Boolean

Public Function SetAsDefaultCursor(ByVal lhWnd As Long, ByVal bState As Boolean)
Dim lIndex As Long
Dim lFirstFreeIndex As Long
Dim i As Long

   If bState Then
      For i = 1 To m_iCount
         If m_tClassCursor(i).hWnd = lhWnd Then
            lIndex = i
            Exit For
         ElseIf m_tClassCursor(i).hWnd = 0 Then
            lFirstFreeIndex = i
         End If
      Next i
      If lIndex = 0 Then
         If lFirstFreeIndex = 0 Then
            m_iCount = m_iCount + 1
            ReDim Preserve m_tClassCursor(1 To m_iCount) As tClassCursor
            lIndex = m_iCount
         Else
            lIndex = lFirstFreeIndex
         End If
      End If
      If m_tClassCursor(lIndex).hCurOld = 0 Then
         m_tClassCursor(lIndex).hWnd = lhWnd
         m_tClassCursor(lIndex).hCurOld = SetClassLong(lhWnd, GCL_HCURSOR, m_hCur)
      End If
   Else
      For i = 1 To m_iCount
         If m_tClassCursor(i).hWnd = lhWnd Then
            lIndex = i
            Exit For
         End If
      Next i
      If lIndex <> 0 Then
         If m_tClassCursor(lIndex).hCurOld <> 0 Then
            SetClassLong lhWnd, GCL_HCURSOR, m_tClassCursor(lIndex).hCurOld
            m_tClassCursor(lIndex).hWnd = 0
            m_tClassCursor(lIndex).hCurOld = 0
         End If
      End If
   End If

End Function


Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
Const CLR_INVALID = -1
    ' Convert Automation color to Windows color
    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Public Function WinAPIError(ByVal lLastDLLError As Long) As String
Dim sBuff As String
Dim lCount As Long
   
   ' Return the error message associated with LastDLLError:
   sBuff = String$(256, 0)
   lCount = FormatMessage( _
      FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
      0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
   If lCount Then
      WinAPIError = Left$(sBuff, lCount)
   End If
   
End Function

Public Property Get hCursor() As Long
   ' Returns the windows handle of the animated
   ' cursor maintained by this class.  Do not
   ' call DeleteObject on this cursor!
   hCursor = m_hCur
End Property

Public Function LoadFromFile(ByVal sFileName As String, Optional ByVal bMapSysColours As Boolean, Optional ByVal bNotAnimated As Boolean = False) As Boolean
Dim lErr As Long
Dim lFlags As Long
   Destroy
   ' Load the animated cursor from the file:
   lFlags = LR_LOADFROMFILE
   ' Set flag for sys colour mapping:
   If bMapSysColours Then
      lFlags = lFlags Or LR_LOADMAP3DCOLORS
   End If
   m_hCur = LoadImageString(App.hInstance, sFileName, IMAGE_CURSOR, 0, 0, lFlags)
   ' Check for errors:
   lErr = Err.LastDllError
   If (m_hCur <> 0) Then
      ' Ok!
      If Not (bNotAnimated) Then
         EvaluateFrames
      End If
      LoadFromFile = True
   Else
      ' Failed.  Return error code:
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
End Function
Public Function LoadFromResource(ByVal lResourceID As Long, Optional ByVal bMapSysColours As Boolean, Optional ByVal hInstance As Long = 0, Optional ByVal bNotAnimated As Boolean = False) As Boolean
Dim lErr As Long
Dim lFlags As Long
   ' Remove existing instance:
   Destroy
   ' Ensure hInstance correct - if missing assume local module:
   If (hInstance = 0) Then
      hInstance = App.hInstance
   End If
   ' Check for sys colour mapping:
   If bMapSysColours Then
      lFlags = lFlags Or LR_LOADMAP3DCOLORS
   End If
   ' Load the animated cursor from the resource:
   m_hCur = LoadImageLong(hInstance, lResourceID, IMAGE_CURSOR, 0, 0, lFlags)
   ' Check for errors:
   lErr = Err.LastDllError
   If (m_hCur <> 0) Then
      ' Ok!
      If Not (bNotAnimated) Then
         EvaluateFrames
      End If
      LoadFromResource = True
   Else
      ' Failed.  Return error code:
      If InDevelopment Then
         Err.Raise 10000, App.EXEName & ".cAniCursor", "Cursors cannot be loaded from resource in the IDE."
      Else
         Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
      End If
   End If
End Function
Private Function InDevelopment() As Boolean
   Debug.Assert pbInDevelopment()
   InDevelopment = m_bInDevelopment
End Function
Private Function pbInDevelopment() As Boolean
   m_bInDevelopment = True
   pbInDevelopment = True
End Function
Private Sub EvaluateFrames()
Dim i As Long
Dim hDCMem As Long
Dim hDCDesk As Long
Dim lR As Long
   ' Determine the number of frames in the animated cursor
   ' by trying to draw it into a memory DC:
   hDCDesk = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   If (hDCDesk <> 0) Then
      hDCMem = CreateCompatibleDC(hDCDesk)
      If hDCMem <> 0 Then
         Do
            lR = DrawIconEx(hDCMem, 0, 0, m_hCur, 0, 0, i, 0, DI_NORMAL Or DI_DEFAULTSIZE)
            If lR = 0 Then
               m_iNumFrames = i
               Exit Do
            Else
               i = i + 1
            End If
         Loop While i < 1024 ' SPM - There must be a smarter way to determine that an image
                             ' has only one frame than this!
         DeleteDC hDCMem
      End If
      DeleteDC hDCDesk
   End If
End Sub

Public Property Get Frame() As Long
   ' Returns the current frame:
   Frame = m_iFrame
End Property
Public Property Let Frame(ByVal iFrame As Long)
   ' Set the frame if in range
   If (iFrame >= 0) And (iFrame < m_iNumFrames) Then
      m_iFrame = iFrame
   Else
      Err.Raise 9, App.EXEName & ".cAniCursor"
   End If
End Property
Public Property Get FrameCount() As Long
   FrameCount = m_iNumFrames
End Property
Public Sub Step()
   m_iFrame = m_iFrame + 1
   If (m_iFrame < 0) Then m_iFrame = 0
   If (m_iFrame >= m_iNumFrames) Then m_iFrame = 0
End Sub
Public Sub Destroy()
Dim i As Long
   If m_hCur <> 0 Then
      For i = 1 To m_iCount
         If m_tClassCursor(i).hCurOld <> 0 Then
            SetClassLong m_tClassCursor(i).hWnd, GCL_HCURSOR, m_tClassCursor(i).hCurOld
         End If
      Next i
      Erase m_tClassCursor
      m_iCount = 0
      DestroyCursor m_hCur
      m_hCur = 0
      m_iNumFrames = 0
      m_iFrame = 0
   End If
End Sub
Public Sub Draw( _
      ByVal hDC As Long, _
      Optional ByVal lLeft As Long = 0, _
      Optional ByVal lTop As Long = 0, _
      Optional ByVal lWidth As Long = 0, _
      Optional ByVal lHeight As Long = 0, _
      Optional ByVal lBackColor As OLE_COLOR = -1 _
   )
' Draw the current frame of the animated cursor on the specified DC.
' optionally you can set left and top position.  If you specify
' width and height then the drawing code will attempt to draw the
' animated cursor at the specified size.
' Setting the backcolor will cause windows to automatically fill the
' background behind the animated cursor with that colour.  The cursor
' will no longer be drawn transparently.

Dim lFlags As Long
Dim hBr As Long
   If (lWidth = 0) Or (lHeight = 0) Then
      lFlags = DI_NORMAL Or DI_DEFAULTSIZE
   Else
      lFlags = DI_NORMAL
   End If
   If lBackColor <> -1 Then
      ' Windows can automatically fill the background
      ' using this brush:
      hBr = CreateSolidBrush(TranslateColor(lBackColor))
   End If
   DrawIconEx hDC, lLeft, lTop, m_hCur, lWidth, lHeight, m_iFrame, hBr, lFlags
   If hBr <> 0 Then
      DeleteObject hBr
   End If
End Sub

Private Sub Class_Terminate()
   ' Ensure we clear up any resources:
   Destroy
End Sub
