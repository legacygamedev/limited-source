Attribute VB_Name = "modAdvMapEditor"
Option Explicit

Private Const GWL_STYLE         As Long = (-16)

Private Const SWP_NOACTIVATE    As Long = &H10
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOZORDER      As Long = &H4
Private Const SWP_NOSIZE        As Long = &H1
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOMOVE        As Long = &H2

'API Declarations
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" _
                        (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Long, lpRect As RECTTT) As Long
Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetWindowPos Lib "user32.dll" ( _
     ByVal hWnd As Long, _
     ByVal hWndInsertAfter As Long, _
     ByVal X As Long, _
     ByVal Y As Long, _
     ByVal cx As Long, _
     ByVal cy As Long, _
     ByVal wFlags As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" _
   Alias "RtlMoveMemory" _
   (hpvDest As Any, _
    hpvSource As Any, _
    ByVal cbCopy As Long)
'Types
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Public Type RECTTT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

'Mod Globals
Private g_MovingMainWnd As Boolean
Private g_OrigCursorPos As POINTAPI
Private g_OrigWndPos As POINTAPI
Private CurrentLayer As String
Public currentMapLayerNum As String
Public displayTilesets As Boolean
Public layersActive As Boolean

Public Function getCurrentMapLayerName() As String
    Dim llayer As OptionButton
    
    For Each llayer In frmEditor_Map.optLayer
        If llayer Then
            currentMapLayerNum = llayer.Index
            CurrentLayer = llayer.Caption
            getCurrentMapLayerName = CurrentLayer
            Exit For
        End If
    Next
    'frmMain.lblTitle = "UBER Map Editor - " & "Layer: " & CurrentLayer
End Function

Public Sub MapEditorMode(switch As Boolean)
    If switch Then
        frmMain.picMapEditor.Top = 0
        frmMain.picMapEditor.Left = 0
        frmMain.picMapEditor.Visible = True
        frmMain.Width = frmMain.Width - 30
        frmMain.Height = frmMain.Height + 1110
        frmMain.picForm.Top = frmMain.picForm.Top + 24 + 50
        
        If frmMain.mapPreviewSwitch.Value Then
            frmMain.mapPreviewSwitch.Picture = LoadResPicture("MAP_DOWN", vbResBitmap)
            frmMapPreview.Show
            frmMapPreview.RecalcuateDimensions
        Else
            frmMain.mapPreviewSwitch.Picture = LoadResPicture("MAP_UP", vbResBitmap)
        End If
        
        'Tile Preview
        If frmMain.chkEyeDropper.Value Then
            frmMain.chkEyeDropper.Picture = LoadResPicture("EYE_DOWN", vbResBitmap)
        Else
            frmMain.chkEyeDropper.Picture = LoadResPicture("EYE_UP", vbResBitmap)
        End If
        
        'Buttons
        frmMain.cmdSave.Picture = LoadResPicture("MAP_SAVE", vbResBitmap)
        frmMain.cmdRevert.Picture = LoadResPicture("MAP_REVERT", vbResBitmap)
        frmMain.cmdDelete.Picture = LoadResPicture("MAP_DELETE", vbResBitmap)
        frmMain.cmdProperties.Picture = LoadResPicture("MAP_PROPERTIES", vbResBitmap)
        frmMain.chkTilesets.Picture = LoadResPicture("TILESETS_UP", vbResBitmap)
        frmMain.chkLayers.Picture = LoadResPicture("LAYERS_UP", vbResBitmap)

        'LAbels
        getCurrentMapLayerName
        EditorSave = True
    Else
        frmMain.picMapEditor.Visible = False
        frmMain.Width = frmMain.Width + 30
        frmMain.Height = frmMain.Height - 1110
        frmMain.picForm.Top = frmMain.picForm.Top - 24 - 50
        Unload frmMapPreview
    End If
    'Call FlipBit(WS_CAPTION, Not switch)
End Sub

Public Sub LeaveMapEditorMode(Cancel As Boolean)
    If EditorSave = False And Cancel Then
         MapEditorCancel
    ElseIf EditorSave = False And Not Cancel Then
        MapEditorLeaveMap
    End If
    
    EditorSave = False
    Call ToggleGUI(True)
    Call ToggleButtons(True)
    
    ' Make sure the properties form is closed
    If FormVisible("frmEditor_MapProperties") Then
        Unload frmEditor_MapProperties
    End If
    
    ' Make sure event editor form is closed
    If FormVisible("frmEditor_Events") Then
        Unload frmEditor_Events
    End If
    
    If FormVisible("frmEditor_Map") Then
        If Not frmEditor_Map.UnloadStarted Then
            frmEditor_Map.UnloadStarted = True
            Unload frmEditor_Map
        End If
        MapEditorMode False
    End If
    
    If FormVisible("frmAdmin") Then
        frmAdmin.ignoreChange = True
        frmAdmin.chkEditor(EDITOR_MAP).Value = 0
        frmAdmin.chkEditor(EDITOR_MAP).FontBold = False
        BringWindowToTop (frmAdmin.hWnd)
        frmAdmin.picEye(EDITOR_MAP).Visible = False
    End If
    
    InMapEditor = False
    
    If Trim$(CurrentMusic) <> Trim$(Map.Music) Then
        PlayMapMusic
    End If
End Sub

Private Function FlipBit(ByVal Bit As Long, ByVal Value As Boolean) As Boolean
   Dim nStyle As Long
   
   nStyle = GetWindowLong(frmMain.hWnd, GWL_STYLE)
   
   If Value Then
      nStyle = nStyle Or Bit
   Else
      nStyle = nStyle And Not Bit
   End If
   Call SetWindowLong(frmMain.hWnd, GWL_STYLE, nStyle)
   Call Redraw
   
   FlipBit = (nStyle = GetWindowLong(frmMain.hWnd, GWL_STYLE))
End Function
Public Sub Redraw()
   ' Redraw window with new style.
   Dim swpFlags As Long
   swpFlags = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOOWNERZORDER
      'SWP_NOZORDER
    SetWindowPos frmMain.hWnd, 0, 0, 0, 0, 0, swpFlags
End Sub

Public Sub MainMouseMove(hWnd As Long)

    If (g_MovingMainWnd) Then

        Dim pt As POINTAPI

        If (GetCursorPos(pt)) Then

            Dim wnd_x As Long, wnd_y As Long

            wnd_x = g_OrigWndPos.X + (pt.X - g_OrigCursorPos.X)
            wnd_y = g_OrigWndPos.Y + (pt.Y - g_OrigCursorPos.Y)
            SetWindowPos frmMain.hWnd, 0, wnd_x, wnd_y, 0, 0, (SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOZORDER Or SWP_NOSIZE)
            If FormVisible("frmMapPreview") Then
                frmMapPreview.Move frmMain.Left - frmMapPreview.Width, frmMain.Top
            End If
        End If
    End If

End Sub
Public Sub MainCaptureChanged(hWnd As Long, lParam As Long)
    g_MovingMainWnd = IIf(lParam = hWnd, True, False)
End Sub
Public Sub MainLButtonUp(hWnd As Long)
    ReleaseCapture
End Sub

Public Sub MainLButtonDown(hWnd As Long)

    If (GetCursorPos(g_OrigCursorPos)) Then

        Dim rt As RECTTT

        GetWindowRect frmMain.hWnd, rt
        g_OrigWndPos.X = rt.Left
        g_OrigWndPos.Y = rt.Top
        g_MovingMainWnd = True
        SetCapture hWnd
    End If

End Sub
Public Sub GetWindowSize(hwndd As Long, ByRef rectt As RECTTT)
        GetWindowRect hwndd, rectt
End Sub
Public Sub BeforeTopMost(hwndd As Long)
    SetWindowPos hwndd, 0, 0, 0, 0, 0, (SWP_NOACTIVATE Or SWP_NOOWNERZORDER Or SWP_NOSIZE Or SWP_NOMOVE)
End Sub
Public Sub MainPreventResizing(hWnd As Long, constWidth As Long, constHeight As Long, ByRef lParam As Long)
                 Dim MMI As MINMAXINFO
                  
                  CopyMemory MMI, ByVal lParam, LenB(MMI)
                   With MMI
                      .ptMinTrackSize.X = constWidth
                      .ptMinTrackSize.Y = constHeight
                      .ptMaxTrackSize.X = constWidth
                      .ptMaxTrackSize.Y = constHeight
                  End With
                  CopyMemory ByVal lParam, MMI, LenB(MMI)
End Sub
